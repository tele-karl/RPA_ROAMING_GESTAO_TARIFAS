from decimal import Decimal, InvalidOperation
import pandas as pd

from src.services.log_service import registrar_log

def converter_unidade(valor):
    """
    Função para converter unidade
    Converte representações comuns para os tokens do loader
    """
    try:
        if not valor or (isinstance(valor, float) and pd.isna(valor)):
            return None
        s = str(valor).strip().lower()
        if "sec" in s:
            return "SECO"
        if "min" in s:
            return "MINUTE"
        if "1 kb" in s:
            return "KBYTE"
        if "10 kb" in s:
            return "10K"
        if "100 kb" in s:
            return "100K"
        if "mb" in s:
            return "MBYTE"
        if "gb" in s:
            return "GBYTE"
        return "UNIT"
    except Exception:
        registrar_log("Finalizado", "Erro", f"Falha ao converter unidade.")
        return None

def formatar_valor(value):
    """
    Formata valores numéricos para:
      - evitar notação científica,
      - remover zeros à direita,
      - remover ponto final se não houver parte decimal,
      - manter '0' para zeros (ex.: 0.0000 -> 0).
    """
    if value is None:
        return ""
    s = str(value).strip()
    if s == "":
        return ""
    try:
        # Aceita vírgula como separador decimal, se aparecer em alguma origem
        s = s.replace(",", ".")
        d = Decimal(s)

        # Gera string em formato fixo (sem expoente)
        f = format(d, "f")  # ex.: '0.02500', '2.000', '0.0000'
        # Remove zeros à direita da parte decimal
        if "." in f:
            f = f.rstrip("0").rstrip(".")   # '0.02500' -> '0.025'; '2.000' -> '2'; '0.0000' -> ''
        # Se esvaziou (caso '0.0000'), retorne '0'
        return f if f != "" else "0"
    except (InvalidOperation, ValueError):
        registrar_log("Finalizado", "Erro", f"Falha ao formatar valor.")
        # Fallback: devolve string original (sem travar o fluxo)
        return str(value)

def formatar_data_spreadsheetml(valor):
    """
    Converte para ISO 8601 requerido por SpreadsheetML quando ss:Type='DateTime':
    'YYYY-MM-DDTHH:MM:SS'. Retorna "" para valores vazios/inválidos.
    """
    try:
        if valor is None or (isinstance(valor, str) and valor.strip() == ""):
            return ""
        # Excel serial -> datetime
        if isinstance(valor, (int, float)) and not isinstance(valor, bool):
            dt = pd.to_datetime(valor, origin="1899-12-30", unit="D")
        else:
            # aceita strings 'dd/mm/yyyy' etc.
            dt = pd.to_datetime(valor, errors="coerce", dayfirst=True)
        if pd.isna(dt):
            return ""
        return dt.strftime("%Y-%m-%dT%H:%M:%S")
    except Exception:
        registrar_log("Finalizado", "Erro", f"Falha ao formatar data (SpreadsheetML).")
        return ""

def formatar_data_ddmmyyyy(valor):
    """
    Formata datas no padrão dd/mm/yyyy para exibição no XLSX/DF.
    Aceita serial do Excel, strings e pandas.Timestamp.
    """
    try:
        if valor is None or (isinstance(valor, str) and valor.strip() == ""):
            return ""
        # Excel serial
        if isinstance(valor, (int, float)) and not isinstance(valor, bool):
            dt = pd.to_datetime(valor, origin="1899-12-30", unit="D")
        else:
            dt = pd.to_datetime(valor, errors="coerce", dayfirst=True)
        if pd.isna(dt):
            return ""
        return dt.strftime("%d/%m/%Y")
    except Exception:
        registrar_log("Finalizado", "Erro", f"Falha ao formatar data (dd-mm-yyyy).")
        return ""