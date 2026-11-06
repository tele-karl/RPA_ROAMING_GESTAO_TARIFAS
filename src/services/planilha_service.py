import openpyxl
from openpyxl.styles import Font, PatternFill
from typing import Dict
from src.services.log_service import registrar_log
from src.utils.excel_utils import (
    aba_base_disc_rates,
    aba_int_call_disc
)

def carregar_pulseforce(filepath: str) -> Dict[str, str]:
    """
    Carrega a planilha Pulseforce Networks e retorna um dicionário
    com Tap Code como chave e Alias como valor.
    """
    try:
        registrar_log("Iniciado", "Sucesso", f"Carregando arquivo Pulseforce: {filepath}")
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active  # Considerando que os dados estão na primeira aba

        tap_to_alias = {}
        # Encontrar índices das colunas Tap Code e Alias
        header_row = 2  # Linha de início dos dados
        tap_code_col = None
        alias_col = None

        for col in range(1, ws.max_column + 1):
            header_value = ws.cell(row=header_row, column=col).value
            if header_value:
                hv = str(header_value).strip().lower()
                if hv == "tap code":
                    tap_code_col = col
                elif hv == "alias":
                    alias_col = col

        if tap_code_col is None or alias_col is None:
            registrar_log("Iniciado", "Erro", "Colunas 'Tap Code' e 'Alias' não encontradas na planilha Pulseforce.")
            wb.close()
            return {}

        # Ler os dados a partir da linha 3
        for row in range(header_row + 1, ws.max_row + 1):
            tap_code = ws.cell(row=row, column=tap_code_col).value
            alias = ws.cell(row=row, column=alias_col).value
            if tap_code and alias:
                tap_to_alias[str(tap_code).strip()] = str(alias).strip()

        registrar_log("Finalizado", "Sucesso", f"Dicionário carregado com {len(tap_to_alias)} entradas.")
        wb.close()
        return tap_to_alias

    except Exception as e:
        registrar_log("Finalizado", "Erro", f"Criação de dicionário com arquivo Pulseforce deu erro: {e}")
        return {}

def inserir_linha_com_alias(ws, row_idx: int, alias: str, preencher_coluna_alias: bool) -> None:
    """
    Duplica a linha abaixo de `row_idx`, porém copia APENAS as colunas de TAP-OUT:
      - Base Disc Rates: 1..28 (30 é a coluna do alias / início do TAP-IN)
      - Int. Call Disc: 1..14 (15 é a coluna do alias / início do TAP-IN)
    Em seguida:
      - substitui a coluna A pelo alias (sempre, TAP-OUT)
      - preenche a coluna do alias (30/15) SOMENTE se preencher_coluna_alias=True
    """
    try:
        print("[INFO]", f"Inserindo linha com alias '{alias}' abaixo da linha {row_idx}")

        # Determina limites TAP-OUT e a coluna de alias conforme a aba
        aba_titulo = ws.title   # não precisa lower aqui; as funções já tratam
        if aba_base_disc_rates(aba_titulo):
            end_col_tapout = 28     # TAP-OUT termina na 28; 30 é a coluna do alias
            alias_col = 30
        elif aba_int_call_disc(aba_titulo):
            end_col_tapout = 14     # TAP-OUT termina na 14; 15 é a coluna do alias
            alias_col = 15
        else:
            print("[ERROR]", f"Aba inválida para inserir alias: {aba_titulo}")
            return

        # Insere uma nova linha abaixo da original
        ws.insert_rows(row_idx + 1)

        # Copia APENAS as colunas de TAP-OUT para a linha do alias
        for col in range(1, end_col_tapout + 1):
            origem_cell = ws.cell(row=row_idx, column=col)
            destino_cell = ws.cell(row=row_idx + 1, column=col)
            destino_cell.value = origem_cell.value
            # Copia estilo básico e marca em negrito
            if origem_cell.font:
                destino_cell.font = Font(
                    name=origem_cell.font.name, 
                    size=origem_cell.font.size, 
                    bold=True
                )
            else:
                destino_cell.font = Font(bold=True)
        # Substitui a coluna A pelo alias (TAP-OUT)
        ws.cell(row=row_idx + 1, column=1).value = alias

        # Preenche a "coluna do alias" (início do TAP-IN) apenas se houver TAP-IN na origem
        if preencher_coluna_alias:
            ws.cell(row=row_idx + 1, column=alias_col).value = alias

    except Exception as e:
        print("[ERROR]", f"Erro ao inserir linha com alias '{alias}' abaixo da linha {row_idx}: {e}")

def inserir_linha_tapin_valores(ws, insert_after_row: int, valores_por_col: dict, col_fill: int) -> None:
    """
    Insere uma nova linha abaixo de insert_after_row e preenche apenas as colunas
    especificadas em valores_por_col (tipicamente 15..max ou 30..max).
    Aplica negrito básico e marca a coluna col_fill em preto.
    """
    try:
        ws.insert_rows(insert_after_row + 1)

        # Escreve apenas as colunas do TAP-IN conforme o dicionário (col_index -> value)
        for c, v in sorted(valores_por_col.items()):
            dst_cell = ws.cell(row=insert_after_row + 1, column=c)
            dst_cell.value = v
            # tenta herdar fonte da linha de cima, se existir; aplica negrito
            up_cell = ws.cell(row=insert_after_row, column=c)
            if up_cell.font:
                dst_cell.font = Font(name=up_cell.font.name, size=up_cell.font.size, bold=True)
            else:
                dst_cell.font = Font(bold=True)

        # marcação visual na coluna de referência da aba
        ws.cell(row=insert_after_row + 1, column=col_fill).fill = PatternFill(
            start_color="000000", end_color="000000", fill_type="solid"
        )
    except Exception as e:
        print("[ERROR]", f"Erro ao inserir linha (TAP-IN): {e}")