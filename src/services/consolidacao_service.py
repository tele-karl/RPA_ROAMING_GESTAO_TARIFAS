import os
import calendar
import openpyxl
from openpyxl.styles import Font, PatternFill
from datetime import datetime
from typing import Dict
from src.configs.configuration import (
    ORIGEM_DIR, DESTINO_DIR, MODELO_PATH,
)
from src.services.log_service import registrar_log
from src.utils.excel_utils import (
    encontrar_aba_por_tipo,
    encontrar_primeira_linha_vazia,
    linha_existe,
    encontrar_primeira_linha_vazia_nas_colunas,
    linha_existe_nas_colunas,
    compactar_parte,
    aba_base_disc_rates,
    aba_int_call_disc
)
from src.services.planilha_service import (
    inserir_linha_com_alias,
    inserir_linha_tapin_valores,
)
from src.services.rede_service import verificar_effective_date

def criar_planilha_consolidada() -> str | None:
    """
    Cria uma nova planilha consolidada com base no arquivo modelo em MODELO_DIR.
    O nome do arquivo é baseado na Effective Date dos arquivos em ORIGEM_DIR com o padrão 'TAP MMM-YY.xlsx'.
    Usa a Effective Date mais recente encontrada entre os arquivos aprovados.
    Se nenhum arquivo em ORIGEM_DIR tiver Effective Date válida, usa 'hoje' como fallback.
    """
    try:
        modelo_abs = os.path.abspath(MODELO_PATH)
        print("[INFO]", f"Criando planilha consolidada... (modelo: {modelo_abs})")
        registrar_log("Iniciado", "Sucesso", f"Criando planilha consolidada (modelo: {modelo_abs})")
        
        # Verificação do modelo
        if not os.path.exists(MODELO_PATH):
            registrar_log("Iniciado", "Erro", f"Arquivo modelo não encontrado: {modelo_abs}")
            return None

        # Coleta Effective Dates dos arquivos em ORIGEM_DIR
        efetivas = []
        os.makedirs(ORIGEM_DIR, exist_ok=True)

        arquivos_origem = [
            f for f in os.listdir(ORIGEM_DIR) 
            if f.lower().endswith(".xlsx") and os.path.isfile(os.path.join(ORIGEM_DIR, f))
        ]

        if not arquivos_origem:
            print("[WARN]", f"Nenhum .xlsx encontrado em {ORIGEM_DIR}. Usando 'hoje' como fallback.")

        for nome in arquivos_origem:
            origem_path = os.path.join(ORIGEM_DIR, nome)
            data_efetiva = verificar_effective_date(origem_path)    # retorna date ou None
            if data_efetiva:
                efetivas.append(data_efetiva)
                print("[INFO]", f"Effective Date '{data_efetiva}' obtida de: {nome}")
            else:
                print("[WARN]", f"Sem Effective Date válida em: {nome}")

        # Escolher a data de referência
        if efetivas:
            referencia_date = max(efetivas)  # mais recente do lote
            print("[INFO]", f"Referência para o nome: {referencia_date} (máxima do lote)")
        else:
            referencia_date = datetime.today().date()
            print("[WARN]", f"Nenhuma Effective Date válida. Usando fallback (hoje): {referencia_date}")

        # Nome do arquivo baseado na referência (forçando inglês p/ padronização)
        nome_abrev_en = calendar.month_abbr[referencia_date.month]  # 'Jan', 'Feb', ...
        nome_arquivo = f"TAP {nome_abrev_en}-{referencia_date.strftime('%y')}.xlsx"

        # Cria arquivo consolidado (cópia do modelo)
        os.makedirs(DESTINO_DIR, exist_ok=True)

        wb_modelo = openpyxl.load_workbook(MODELO_PATH)
        destino_final = os.path.join(DESTINO_DIR, nome_arquivo)
        wb_modelo.save(destino_final)
        wb_modelo.close()
        
        registrar_log("Finalizado", "Sucesso", f"Planilha consolidada criada: {os.path.abspath(destino_final)}")
        return destino_final
    
    except Exception as e:
        registrar_log("Finalizado", "Erro", f"Planilha consolidada não foi criada: {e}")

def limpar_linhas_vazias_por_parte(wb) -> None:
    """
    Executa a limpeza por PARTES, em ordem:
      1) Base Disc Rates -> TAP-OUT (1..28) a partir da linha 16
      2) Base Disc Rates -> TAP-IN  (30..máx) a partir da linha 16
      3) Int. Call Disc  -> TAP-OUT (1..13) a partir da linha 12
      4) Int. Call Disc  -> TAP-IN  (15..máx) a partir da linha 12
    """
    try:
        for sheet_name in wb.sheetnames:
            if aba_base_disc_rates(sheet_name):
                ws = wb[sheet_name]
                # TAP-OUT (1..28)
                compactar_parte(ws, start_row=16, col_start=1, col_end=28)
                # TAP-IN (30..max)
                compactar_parte(ws, start_row=16, col_start=30, col_end=ws.max_column)
                
            elif aba_int_call_disc(sheet_name):
                ws = wb[sheet_name]
                # TAP-OUT (1..13)
                compactar_parte(ws, start_row=12, col_start=1, col_end=13)
                # TAP-IN (15..max)
                compactar_parte(ws, start_row=12, col_start=15, col_end=ws.max_column)
        print("[INFO]", f"Limpeza por partes concluída e arquivo final salvo.")
    except Exception as e:
        print("[ERROR]", f"Limpeza por partes da planilha consolidada não realizada: {e}")

def processar_planilhas(tap_to_alias: Dict[str, str]) -> None:
    """
    Processa arquivos de ORIGEM_DIR e insere dados na planilha consolidada (DESTINO_DIR).
    Regras implementadas:
      - Quando não houver TAP-IN na origem, não preenche a coluna do alias (30/15).
      - Para 'Int. Call Disc' com G12 falso: processa somente TAP-IN.
      - No TAP-OUT habilitado: a linha de alias TAP-OUT é criada e, se houver TAP-IN, o TAP-IN do alias
        é replicado na MESMA linha do alias (evitando a linha isolada do alias).
      - Ao final de cada arquivo de destino, executa limpeza por PARTES (TAP-OUT e TAP-IN) em ambas as abas.
    """
    try:
        # Iterar sobre todos os arquivos de destino (colocado inicialmente por segunrança)
        for destino_file in os.listdir(DESTINO_DIR):
            if destino_file.lower().endswith(".xlsx"):
                destino_path = os.path.join(DESTINO_DIR, destino_file)
                registrar_log("Iniciado", "Sucesso", f"Processando planilha consolidada: {destino_path}")
                wb_destino = openpyxl.load_workbook(destino_path)

                # Iterar sobre todos os arquivos de origem
                for origem_file in os.listdir(ORIGEM_DIR):
                    if origem_file.lower().endswith(".xlsx"):
                        origem_path = os.path.join(ORIGEM_DIR, origem_file)
                        print("[INFO]", f"Copiando dados do arquivo de origem: {origem_path}")
                        try:
                            wb_origem = openpyxl.load_workbook(origem_path, data_only=True)
                        except Exception as e:
                            registrar_log("Iniciado", "Erro", f"Falha ao abrir arquivo de origem {origem_path}: {e}")
                            continue

                        for tipo, start_row in [("Base Disc Rates", 16), ("Int. Call Disc", 12)]:
                            aba_origem_nome = encontrar_aba_por_tipo(wb_origem, tipo, log_info=True, label="ORIGEM")
                            aba_destino_nome = encontrar_aba_por_tipo(wb_destino, tipo, log_info=True, label="DESTINO")
                            if aba_origem_nome and aba_destino_nome:
                                ws_origem = wb_origem[aba_origem_nome]
                                ws_destino = wb_destino[aba_destino_nome]

                                # Limites por aba
                                tap_in_start_col = 30 if tipo == "Base Disc Rates" else 15
                                col_fill = 29 if tipo == "Base Disc Rates" else 14

                                # (A) Regras específicas para "Int. Call Disc": G12 falso => sem TAP-OUT
                                processar_tap_out = True
                                processar_tap_in = True
                                if tipo == "Int. Call Disc":
                                    valor_g12 = ws_origem.cell(row=12, column=7).value  # G12
                                    # Normalização segura para string (remove espaços e padroniza caixa)
                                    if isinstance(valor_g12, str):
                                        g12_norm = valor_g12.strip().upper()
                                    else:
                                        g12_norm = valor_g12  # mantém como está (None, bool, número etc.)
                                    if g12_norm in [None, False, "FALSE", "AA14"]:
                                        processar_tap_out = False   # desliga TAP-OUT
                                        processar_tap_in = True     # mantém TAP-IN
                                        print("[WARN]", "IDD Rate falso. TAP-OUT desativado; TAP-IN será processado se houver dados.")

                                # (B) Detectar se HÁ TAP-IN na origem
                                tem_tap_in = False
                                for row_idx in range(start_row, ws_origem.max_row + 1):
                                    v = ws_origem.cell(row=row_idx, column=tap_in_start_col).value  # tap_in_start_col: 30 ou 15, conforme a aba
                                    # Se houver PLMN ou qualquer valor significativo na coluna de TAP-IN,
                                    # significa que a parte TAP-IN da tabela realmente existe.
                                    if v not in (None, "", " "):
                                        tem_tap_in = True
                                        break

                                # PROCESSAR TAP-OUT (se habilitado)
                                if processar_tap_out:
                                    # Largura efetiva da cópia da linha base:
                                    # - Se tem TAP-IN, mantém cópia completa (traz TAP-OUT + TAP-IN na linha base)
                                    # - Se não tem TAP-IN, copia apenas TAP-OUT
                                    max_col_effective = ws_origem.max_column if tem_tap_in else (tap_in_start_col - 1)

                                    linha_insercao = encontrar_primeira_linha_vazia(ws_destino, start_row, max_col_effective)

                                    for row in ws_origem.iter_rows(min_row=start_row, 
                                                                max_row=ws_origem.max_row, 
                                                                max_col=max_col_effective):
                                        valores_origem = [cell.value for cell in row]
                                        if not any(v is not None for v in valores_origem):
                                            continue

                                        if not linha_existe(ws_destino, valores_origem, start_row, max_col_effective):
                                            # 1) Copia a linha base
                                            for c_idx, cell in enumerate(row, start=1):
                                                dest_cell = ws_destino.cell(row=linha_insercao, column=c_idx)
                                                dest_cell.value = cell.value
                                                if cell.font:
                                                    dest_cell.font = Font(name=cell.font.name, size=cell.font.size)
                                            
                                            # Marcação visual (linha base)
                                            ws_destino.cell(row=linha_insercao, column=col_fill).fill = PatternFill(
                                                start_color="000000", end_color="000000", fill_type="solid"
                                            )
                                            
                                            # 2) Aliases do PLMN (coluna A)
                                            plmn = row[0].value
                                            plmn_str = str(plmn).strip() if plmn else None

                                            if plmn_str and plmn_str in tap_to_alias:
                                                aliases = tap_to_alias[plmn_str]
                                                if isinstance(aliases, str):
                                                    aliases = [aliases]

                                                for alias in aliases:
                                                    # 2.a) Inserir linha de alias TAP-OUT imediatamente abaixo da base
                                                    inserir_linha_com_alias(
                                                        ws_destino, linha_insercao,
                                                        alias, preencher_coluna_alias=tem_tap_in,   # mantém a regra (só escreve 30/15 se houver TAP-IN)
                                                    )
                                                    alias_row = linha_insercao + 1

                                                    # marcação na linha do alias TAP-OUT
                                                    ws_destino.cell(row=alias_row, column=col_fill).fill = PatternFill(
                                                        start_color="000000", end_color="000000", fill_type="solid"
                                                    )

                                                    # 2.b) Em vez de criar NOVA linha TAP-IN do alias,
                                                    #      preenche TAP-IN na MESMA linha do alias (alias_row)
                                                    if tem_tap_in:
                                                        origem_row_idx = row[0].row
                                                        # Para Int. Call Disc, ignorar TAP-IN desta linha se IDD Rate estiver vazio.
                                                        if tipo == "Int. Call Disc":
                                                            idd_rate_val = ws_origem.cell(row=origem_row_idx, column=21).value
                                                            if idd_rate_val in (None, "", " "):
                                                                print("[INFO]", f"Ignorando TAP-IN do alias na row {origem_row_idx} por IDD Rate vazio.")
                                                                # Não copiar TAP-IN para esta linha do alias
                                                                continue
                                                        for c in range(tap_in_start_col, ws_origem.max_column + 1):
                                                            src_cell = ws_origem.cell(row=origem_row_idx, column=c)
                                                            dst_cell = ws_destino.cell(row=alias_row, column=c)
                                                            
                                                            if c == tap_in_start_col:
                                                                dst_cell.value = alias  # coluna do alias
                                                            else:
                                                                dst_cell.value = src_cell.value

                                                            if src_cell.font:
                                                                dst_cell.font = Font(name=src_cell.font.name, size=src_cell.font.size, bold=True)
                                                            else:
                                                                dst_cell.font = Font(bold=True)

                                                    # avança para a próxima posição após base + alias
                                                    linha_insercao = alias_row + 1
                                            else:
                                                # sem alias para o PLMN, apenas a base
                                                linha_insercao += 1

                                    # Salva após processar cada aba            
                                    wb_destino.save(destino_path)
                                    registrar_log("Finalizado", "Sucesso", f"Arquivo {destino_path} atualizado com inserção de linhas de {origem_path} (TAP-OUT)")

                                # PROCESSAR SOMENTE TAP-IN (se for o caso)
                                if processar_tap_in and not processar_tap_out:
                                    if not tem_tap_in:
                                        print("[WARN]", "Aba 'Int. Call Disc' sem dados TAP-IN.")
                                    else:
                                        min_col_tapin = tap_in_start_col
                                        max_col_tapin = ws_origem.max_column
                                        col_indices = list(range(min_col_tapin, max_col_tapin + 1))

                                        linha_insercao = encontrar_primeira_linha_vazia_nas_colunas(ws_destino, start_row, col_indices)

                                        for r in range(start_row, ws_origem.max_row + 1):
                                            # valores originais de TAP-IN
                                            valores_tapin = {}
                                            tem_valor_na_linha = False
                                            for c in col_indices:
                                                v = ws_origem.cell(row=r, column=c).value
                                                valores_tapin[c] = v
                                                if v is not None:
                                                    tem_valor_na_linha = True
                                            if not tem_valor_na_linha:
                                                continue
                                            
                                            # Ignorar linhas cujo IDD Rate (coluna U = 21) esteja vazio na parte TAP-IN.
                                            if tipo == "Int. Call Disc":
                                                idd_rate_val = ws_origem.cell(row=r, column=21).value
                                                if idd_rate_val in (None, "", " "):
                                                    print("[INFO]", f"Ignorando linha TAP-IN row {r} por IDD Rate (coluna U) vazio.")
                                                    continue

                                            # 1) Inserir a linha base (TAP-IN original), se não existir
                                            if not linha_existe_nas_colunas(ws_destino, valores_tapin, start_row, col_indices):
                                                # escreve diretamente no destino (apenas colunas TAP-IN)
                                                for c in col_indices:
                                                    src_cell = ws_origem.cell(row=r, column=c)
                                                    dst_cell = ws_destino.cell(row=linha_insercao, column=c)
                                                    dst_cell.value = src_cell.value
                                                    if src_cell.font:
                                                        dst_cell.font = Font(name=src_cell.font.name, size=src_cell.font.size)
                                                # marcação (coluna 14 para Int. Call Disc)
                                                ws_destino.cell(row=linha_insercao, column=14).fill = PatternFill(
                                                    start_color="000000", end_color="000000", fill_type="solid"
                                                )

                                                # 2) Aliases de TAP-IN com base no PLMN da COLUNA 15 (início do TAP-IN)
                                                plmn_tapin = ws_origem.cell(row=r, column=min_col_tapin).value
                                                plmn_tapin_str = str(plmn_tapin).strip() if plmn_tapin else None
                                                if plmn_tapin_str and plmn_tapin_str in tap_to_alias:
                                                    aliases = tap_to_alias[plmn_tapin_str]
                                                    if isinstance(aliases, str):
                                                        aliases = [aliases]

                                                    for alias in aliases:
                                                        valores_tapin_alias = dict(valores_tapin)
                                                        valores_tapin_alias[min_col_tapin] = alias
                                                        # Dedup apenas TAP-IN
                                                        if not linha_existe_nas_colunas(ws_destino, valores_tapin_alias, start_row, col_indices):
                                                            inserir_linha_tapin_valores(
                                                                ws_destino, 
                                                                insert_after_row=linha_insercao,  # abaixo da base recém inserida
                                                                valores_por_col=valores_tapin_alias, 
                                                                col_fill=14
                                                            )
                                                            linha_insercao += 2  # base + alias
                                                        else:
                                                            linha_insercao += 1  # só a base
                                                else:
                                                    linha_insercao += 1  # só a base
                                            else:
                                                # já existia a base; ainda assim inserir os aliases (cada um em nova linha)
                                                plmn_tapin = ws_origem.cell(row=r, column=min_col_tapin).value
                                                plmn_tapin_str = str(plmn_tapin).strip() if plmn_tapin else None
                                                if plmn_tapin_str and plmn_tapin_str in tap_to_alias:
                                                    aliases = tap_to_alias[plmn_tapin_str]
                                                    if isinstance(aliases, str):
                                                        aliases = [aliases]
                                                    for alias in aliases:
                                                        valores_tapin_alias = dict(valores_tapin)
                                                        valores_tapin_alias[min_col_tapin] = alias
                                                        if not linha_existe_nas_colunas(ws_destino, valores_tapin_alias, start_row, col_indices):
                                                            inserir_linha_tapin_valores(
                                                                ws_destino, 
                                                                insert_after_row=linha_insercao,
                                                                valores_por_col=valores_tapin_alias, 
                                                                col_fill=14
                                                            )
                                                            linha_insercao += 1

                                        # Salva após processar a aba                        
                                        wb_destino.save(destino_path)
                                        registrar_log("Finalizado", "Sucesso", f"Arquivo {destino_path} atualizado com inserção de linhas de {origem_path} (TAP-IN)")

                # LIMPEZA POR PARTES EM TODAS AS ABAS DO DESTINO
                limpar_linhas_vazias_por_parte(wb_destino)
                wb_destino.save(destino_path)
                wb_destino.close()

            registrar_log("Finalizado", "Sucesso", f"Processamento da planilha consolidada concluído: {destino_path}")

    except Exception as e:
        registrar_log("Finalizado", "Erro", f"Houve falha ao processar as planilhas: {e}")