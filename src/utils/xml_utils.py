import os
from lxml import etree
import copy
import pandas as pd
from src.utils.format_utils import formatar_data_spreadsheetml
from src.services.log_service import registrar_log

def rows_to_matrix(rows, ns):
    """
    Converte uma lista de elementos <Row> em matriz (lista de listas),
    respeitando ss:Index (células esparsas) e preenchendo para o mesmo número de colunas.
    Obs: assume linhas densas como no XML vindo do Excel)
    """
    matrix = []
    max_cols = 0
    for row in rows:
        cells = row.findall('ss:Cell', ns)
        # Monta linha respeitando ss:Index (base 1)
        linha = []
        col_ptr = 1
        for cell in cells:
            idx_attr = cell.get('{urn:schemas-microsoft-com:office:spreadsheet}Index')
            if idx_attr:
                idx = int(idx_attr)
                # Preenche vazios até a coluna idx-1
                while col_ptr < idx:
                    linha.append("")
                    col_ptr += 1
            data = cell.find('ss:Data', ns)
            linha.append(data.text if data is not None else "")
            col_ptr += 1
        max_cols = max(max_cols, len(linha))
        matrix.append(linha)
    # padding
    matrix = [r + [""] * (max_cols - len(r)) for r in matrix]
    return matrix

def carregar_xml_para_dfs(caminho_xml):
    """
    Lê um XML Excel 2003 (SpreadsheetML) e retorna um dicionário por sheet:
    {
      sheet_name: {
         "df": DataFrame (dados a partir da 3ª linha),
         "estrutura": list (1ª linha),
         "cabecalho": list (2ª linha)
      },
      ...
    }
    Também retorna (tree, root, ns) para reutilização na atualização final do XML.
    """
    try:
        registrar_log("Iniciado", "Sucesso", f"Carregando dados do XML: {caminho_xml}")
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(caminho_xml, parser)
        root = tree.getroot()
        ns = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}

        planilhas = {}
        worksheets = root.xpath(".//ss:Worksheet", namespaces=ns)
        if not worksheets:
            registrar_log("Iniciado", "Erro", "Nenhuma Worksheet encontrada no XML.")

        for ws in worksheets:
            sheet_name = ws.get("{urn:schemas-microsoft-com:office:spreadsheet}Name") or "Sheet1"
            rows = ws.xpath(".//ss:Row", namespaces=ns)
            if not rows:
                registrar_log("Iniciado", "Erro", f"Worksheet '{sheet_name}' sem linhas.")
                planilhas[sheet_name] = {
                    "df": pd.DataFrame(),
                    "estrutura": [],
                    "cabecalho": []
                }
                continue

            matrix = rows_to_matrix(rows, ns)
            linha_estrutura = matrix[0] if len(matrix) > 0 else []
            linha_cabecalho  = matrix[1] if len(matrix) > 1 else []
            dados = matrix[2:] if len(matrix) > 2 else []
            df = pd.DataFrame(dados)

            planilhas[sheet_name] = {
                "df": df,
                "estrutura": linha_estrutura,
                "cabecalho": linha_cabecalho
            }
        registrar_log("Finalizado", "Sucesso", "Dados do XML carregados com sucesso.")
        return planilhas, tree, root, ns
    except Exception as e:
        registrar_log("Finalizado", "Erro", f"Erro ao carregar XML: {e}")

def salvar_xml_excel(tree, caminho_xml_destino):
    """
    Salva o XML SpreadsheetML garantindo:
      - Declaração XML na 1ª linha;
      - Processing Instruction do Excel ('<?mso-application progid="Excel.Sheet"?>') na 2ª linha;
      - Preserva namespaces/estilos usando apenas lxml.
    Pré-condição: 'tree' é um lxml.etree._ElementTree.
    """
    # Garante pasta destino
    os.makedirs(os.path.dirname(caminho_xml_destino), exist_ok=True)
    # Obtém root e parent; em alguns XMLs o root não tem parent (árvore raiz "pura")
    root = tree.getroot()
    parent = root.getparent()
    # Cria um wrapper temporário apenas se precisar inserir a PI "como irmã" do root
    created_wrapper = False
    if parent is None:
        parent = etree.Element('root_wrapper')
        parent.append(root)
        created_wrapper = True
    # Remove qualquer PI duplicada existente com target "mso-application"
    for node in list(parent):
        if isinstance(node, etree._ProcessingInstruction) and node.target == 'mso-application':
            parent.remove(node)
    # Cria PI do Excel e insere antes do root (2ª linha na serialização)
    pi = etree.ProcessingInstruction('mso-application', 'progid="Excel.Sheet"')
    parent.insert(0, pi)
    # Serialização:
    # - Se usar wrapper, precisa escrever manualmente: declaração XML + PI + root
    # - Caso contrário, pode serializar normalmente; o lxml preservará a PI inserida como "irmã" do root
    if created_wrapper:
        with open(caminho_xml_destino, 'wb') as f:
            f.write(b'<?xml version="1.0" encoding="utf-8"?>\n')
            f.write(etree.tostring(pi) + b"\n")
            # etree.tostring(root) já serializa com os namespaces existentes no documento
            f.write(etree.tostring(root, encoding='utf-8'))
    else:
        # Serializa a árvore original com a declaração XML;
        # a PI foi inserida no parent do root e será preservada
        tree.write(caminho_xml_destino, encoding='utf-8', xml_declaration=True)

def atualizar_xml_multiplas_abas(
    tree, root, ns, caminho_xml_destino, planilhas_finais,
    sobrescrever_formulas=False):
    """
    Reconstrói completamente o <Table> com base no DF final por aba,
    preservando: <Column> (larguras), estilos das 2 primeiras linhas, filtros/freeze panes.
    Corrige AutoFilter para iniciar na linha 2 (cabeçalho).
    """
    try:
        total_atualizadas = 0    

        worksheets = root.findall('.//ss:Worksheet', ns)
        if not worksheets:
            registrar_log("Finalizado", "Erro", "Nenhuma Worksheet encontrada no XML; salvando cópia sem alterações.")
            # Ainda assim salva o XML original (para diagnóstico)
            salvar_xml_excel(tree, caminho_xml_destino)
            return
        xns = {'x': 'urn:schemas-microsoft-com:office:excel'}

        for nome, pack in planilhas_finais.items():
            df = pack.get('df') if isinstance(pack, dict) else pack
            if df is None:
                continue
            # localizar Worksheet por Name
            ws = None
            for w in worksheets:
                wname = w.get('{urn:schemas-microsoft-com:office:spreadsheet}Name') or 'Sheet1'
                if wname == nome:
                    ws = w
                    break
            if ws is None:
                registrar_log("Iniciado", "Erro", f"Aba '{nome}' não encontrada no XML; pulando.")
                continue

            table = ws.find('ss:Table', ns)
            if table is None:
                registrar_log("Iniciado", "Erro", f"Worksheet '{nome}' sem <Table>; pulando.")
                continue

            # Guardar <Column>
            columns = [child for child in list(table) if child.tag.endswith('Column')]

            # Primeiras linhas e 1ª linha de dados do XML original para estilos
            all_rows = [child for child in list(table) if child.tag.endswith('Row')]
            row_struct = copy.deepcopy(all_rows[0]) if len(all_rows) >= 1 else None
            row_header = copy.deepcopy(all_rows[1]) if len(all_rows) >= 2 else None
            row_firstdata = all_rows[2] if len(all_rows) >= 3 else None

            percol_style, percol_type = {}, {}
            if row_firstdata is not None:
                def iter_cells_with_index(row_elem):
                    col = 1
                    for c in row_elem.findall('ss:Cell', ns):
                        idx = c.get('{urn:schemas-microsoft-com:office:spreadsheet}Index')
                        if idx:
                            col = int(idx)
                        yield col, c
                        col += 1
                for cidx, cell in iter_cells_with_index(row_firstdata):
                    sid = cell.get('{urn:schemas-microsoft-com:office:spreadsheet}StyleID')
                    percol_style[cidx] = sid
                    data = cell.find('ss:Data', ns)
                    if data is not None:
                        dtype = data.get('{urn:schemas-microsoft-com:office:spreadsheet}Type')
                        percol_type[cidx] = dtype

            # Limpa todas as <Row> "antigas"
            for r in list(table):
                if r.tag.endswith('Row'):
                    table.remove(r)
            # Reinsere <Column> preservadas
            for c in list(table):
                if c.tag.endswith('Column'):
                    table.remove(c)
            for c in columns:
                table.append(c)

            # Helpers de construção de linha
            def build_row_from_values(values, base_row):
                if base_row is None:
                    row_elem = etree.Element('{urn:schemas-microsoft-com:office:spreadsheet}Row')
                    for v in values:
                        cell = etree.Element('{urn:schemas-microsoft-com:office:spreadsheet}Cell')
                        data = etree.Element('{urn:schemas-microsoft-com:office:spreadsheet}Data')
                        data.set('{urn:schemas-microsoft-com:office:spreadsheet}Type', 'String')
                        data.text = '' if (pd.isna(v) or str(v).strip().lower() == 'nan') else str(v)
                        cell.append(data)
                        row_elem.append(cell)
                    # Remover fórmulas em todas as linhas quando solicitado
                    if sobrescrever_formulas:
                        formula_attr = '{urn:schemas-microsoft-com:office:spreadsheet}Formula'
                        for c in row_elem.findall('ss:Cell', ns):
                            if formula_attr in c.attrib:
                                del c.attrib[formula_attr]
                    return row_elem
                row_elem = copy.deepcopy(base_row)
                col = 1
                for c in row_elem.findall('ss:Cell', ns):
                    idx_attr = c.get('{urn:schemas-microsoft-com:office:spreadsheet}Index')
                    if idx_attr:
                        col = int(idx_attr)
                    val = values[col-1] if col-1 < len(values) else ''
                    d = c.find('ss:Data', ns)
                    if d is None:
                        d = etree.Element('{urn:schemas-microsoft-com:office:spreadsheet}Data')
                        d.set('{urn:schemas-microsoft-com:office:spreadsheet}Type', 'String')
                        c.append(d)
                    d.text = '' if (pd.isna(val) or str(val).strip().lower() == 'nan') else str(val)
                    col += 1
                # Remover fórmulas quando solicitado
                if sobrescrever_formulas:
                    formula_attr = '{urn:schemas-microsoft-com:office:spreadsheet}Formula'
                    for c in row_elem.findall('ss:Cell', ns):
                        if formula_attr in c.attrib:
                            del c.attrib[formula_attr]
                return row_elem

            # Estrutura e Cabeçalho
            if len(df) >= 1:
                table.append(build_row_from_values(list(df.iloc[0]), row_struct))
            if len(df) >= 2:
                table.append(build_row_from_values(list(df.iloc[1]), row_header))

            # Dados
            def to_iso_if_datetime(val):
                try:
                    if val is None or (isinstance(val, str) and val.strip() == ''):
                        return ''
                    return formatar_data_spreadsheetml(val)
                except Exception:
                    return ''
            n_rows = len(df)
            n_cols = len(df.columns) if n_rows > 0 else 0
            for ridx in range(2, n_rows):
                vals = list(df.iloc[ridx])
                row_elem_new = etree.Element('{urn:schemas-microsoft-com:office:spreadsheet}Row')
                for j, v in enumerate(vals, start=1):
                    cell = etree.Element('{urn:schemas-microsoft-com:office:spreadsheet}Cell')
                    sid = percol_style.get(j)
                    if sid:
                        cell.set('{urn:schemas-microsoft-com:office:spreadsheet}StyleID', sid)
                    data = etree.Element('{urn:schemas-microsoft-com:office:spreadsheet}Data')
                    dtype = percol_type.get(j, 'String')
                    if dtype == 'DateTime':
                        if (pd.isna(v) or str(v).strip() == ''):
                            row_elem_new.append(cell)
                            continue
                        data.set('{urn:schemas-microsoft-com:office:spreadsheet}Type', 'DateTime')
                        data.text = to_iso_if_datetime(v)
                    else:
                        data.set('{urn:schemas-microsoft-com:office:spreadsheet}Type', 'String')
                        data.text = '' if (pd.isna(v) or str(v).strip().lower() == 'nan') else str(v)
                    cell.append(data)
                    row_elem_new.append(cell)
                # Remover fórmulas quando solicitado
                if sobrescrever_formulas:
                    formula_attr = '{urn:schemas-microsoft-com:office:spreadsheet}Formula'
                    for c in row_elem_new.findall('ss:Cell', ns):
                        if formula_attr in c.attrib:
                            del c.attrib[formula_attr]
                table.append(row_elem_new)

            # Atualiza contagens e AutoFilter (linha 2 como header)
            if n_cols:
                table.set('{urn:schemas-microsoft-com:office:spreadsheet}ExpandedColumnCount', str(n_cols))
            rebuilt_rows = sum(1 for ch in table if ch.tag.endswith('Row'))
            table.set('{urn:schemas-microsoft-com:office:spreadsheet}ExpandedRowCount', str(rebuilt_rows))
            try:
                af = ws.find('.//x:AutoFilter', xns)
                if af is not None and n_cols:
                    new_range = f"R2C1:R{rebuilt_rows}C{n_cols}"
                    af.set('{urn:schemas-microsoft-com:office:excel}Range', new_range)
            except Exception:
                pass

            total_atualizadas += 1

        salvar_xml_excel(tree, caminho_xml_destino)
        registrar_log("Finalizado", "Sucesso", f"XML final (preservando formatação) salvo em: {caminho_xml_destino} (abas atualizadas: {total_atualizadas})")
        print(f"XML final (preservando formatação) salvo em: {caminho_xml_destino} (abas atualizadas: {total_atualizadas})")

    except Exception as e:
        registrar_log("Finalizado", "Erro", f"Erro na atualização do XML (preservando formatação): {e}")
        print(f"[ERRO] Falha na atualização do XML (preservando formatação): {e}")