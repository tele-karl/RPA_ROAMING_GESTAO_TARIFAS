from datetime import datetime, date
from typing import Optional
from openpyxl.styles import Font, PatternFill
from src.services.log_service import registrar_log

# Identificação de abas
def aba_base_disc_rates(sheet_name: str) -> bool:
    """
    Retorna True se o nome da aba representar 'Base Disc Rates'
    em qualquer uma das variações conhecidas (case-insensitive).
    """
    s = sheet_name.lower()
    return ("brasil base disc rates" in s) or ("vivo base disc rates" in s)

def aba_int_call_disc(sheet_name: str) -> bool:
    """
    Retorna True se o nome da aba representar 'Int. Call Disc'
    em qualquer uma das variações conhecidas (case-insensitive).
    """
    s = sheet_name.lower()
    return ("brasil int. call disc" in s) or ("vivo int. call disc" in s)

def encontrar_aba_por_tipo(wb, tipo: str, log_info: bool = False, label: str = "") -> Optional[str]:
    """
    Encontra a aba no workbook que contém o tipo especificado.
    tipo: "Base Disc Rates" ou "Int. Call Disc"
    Reconhece variantes como 'Brasil ...' e 'VIVO ...' (case-insensitive).
    """
    tipo_lc = tipo.lower()
    for sheet in wb.sheetnames:
        if tipo_lc == "base disc rates" and aba_base_disc_rates(sheet):
            if log_info:
                print("[INFO]", f"Aba {label} ({tipo}) encontrada: {sheet}")
            return sheet
        if tipo_lc == "int. call disc" and aba_int_call_disc(sheet):
            if log_info:
                print("[INFO]", f"Aba {label} ({tipo}) encontrada: {sheet}")
            return sheet
    registrar_log("Finalizado", "Erro", f"Nenhuma aba encontrada para tipo: {tipo}")
    return None

# Datas
def parse_date_generic(val) -> Optional[date]:
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(val.strip(), fmt).date()
            except ValueError:
                pass
    return None

# Linhas/colunas
def encontrar_primeira_linha_vazia(ws, start_row: int, max_col: int) -> int:
    """
    Encontra a primeira linha vazia a partir de start_row, avaliando colunas 1..max_col.
    """
    for r in range(start_row, ws.max_row + 2):
        if all(ws.cell(row=r, column=c).value is None for c in range(1, max_col + 1)):
            return r
    return ws.max_row + 1

def linha_existe(ws, valores: list, start_row: int, max_col: int) -> bool:
    """
    Verifica se uma linha com os mesmos valores já existe no destino.
    Compara todas as colunas TAP-OUT/TAP-IN (1..max_col) fornecidas.
    """
    for dest_row in range(start_row, ws.max_row + 1):
        if all(ws.cell(row=dest_row, column=c_idx).value == valores[c_idx - 1] for c_idx in range(1, max_col + 1)):
            return True
    return False

def linha_existe_nas_colunas(ws, valores_por_coluna: dict, start_row: int, col_indices: list[int]) -> bool:
    """
    Verifica se já existe uma linha no destino que tenha os mesmos valores
    nas colunas especificadas em col_indices.
    valores_por_coluna: dict {col_index: value} apenas para as colunas relevantes.
    """
    for dest_row in range(start_row, ws.max_row + 1):
        iguais = True
        for c in col_indices:
            if ws.cell(row=dest_row, column=c).value != valores_por_coluna.get(c):
                iguais = False
                break
        if iguais:
            return True
    return False

def encontrar_primeira_linha_vazia_nas_colunas(ws, start_row: int, col_indices: list[int]) -> int:
    """
    Encontra a primeira linha vazia a partir de start_row, considerando
    como 'vazia' a linha em que TODAS as colunas de col_indices são None.
    """
    for r in range(start_row, ws.max_row + 2):
        if all(ws.cell(row=r, column=c).value is None for c in col_indices):
            return r
    return ws.max_row + 1

# Limpeza por partes
def intervalo_linha_vazio(ws, row: int, col_start: int, col_end: int) -> bool:
    """
    Retorna True se TODAS as células [col_start..col_end] da 'row' estiverem vazias (None ou string vazia).
    """
    for c in range(col_start, col_end + 1):
        val = ws.cell(row=row, column=c).value
        if val not in (None, ""):
            return False
    return True

def compactar_parte(ws, start_row: int, col_start: int, col_end: int, clear_styles: bool = True) -> None:
    """
    'Compacta' uma PARTE (faixa de colunas) da aba:
      - varre da linha 'start_row' até o final;
      - 'move para cima' apenas as linhas cujo intervalo [col_start..col_end] NÃO esteja vazio;
      - limpa (None) o intervalo [col_start..col_end] das linhas remanescentes no rodapé.
    NÃO deleta linhas da planilha; atua somente no intervalo de colunas informado.
    """
    write_row = start_row
    max_row = ws.max_row

    # 1) Move "para cima" as linhas com dados na faixa [col_start..col_end]
    for read_row in range(start_row, max_row + 1):
        if not intervalo_linha_vazio(ws, read_row, col_start, col_end):
            if read_row != write_row:
                for c in range(col_start, col_end + 1):
                    src = ws.cell(row=read_row, column=c)
                    dst = ws.cell(row=write_row, column=c)
                    # move valor
                    dst.value = src.value
                    # preserva estilo básico
                    if src.font:
                        dst.font = Font(name=src.font.name, size=src.font.size, bold=src.font.bold)
                    if src.fill and src.fill.fill_type:
                        dst.fill = PatternFill(
                            start_color=(src.fill.start_color.rgb if hasattr(src.fill.start_color, 'rgb') else "000000"),
                            end_color=(src.fill.end_color.rgb if hasattr(src.fill.end_color, 'rgb') else "000000"),
                            fill_type=src.fill.fill_type
                        )
                    # limpa origem
                    src.value = None
                    if clear_styles:
                        src.font = Font()
                        src.fill = PatternFill()
            write_row += 1
            
    # 2) Limpa o intervalo restante (rodapé da parte)       
    for r in range(write_row, max_row + 1):
        for c in range(col_start, col_end + 1):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            if clear_styles:
                cell.font = Font()
                cell.fill = PatternFill()