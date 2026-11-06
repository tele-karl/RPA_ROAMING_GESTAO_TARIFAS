import os
from typing import Optional, Tuple, List, Dict
from decimal import Decimal, InvalidOperation
from datetime import datetime, date
import openpyxl
from openpyxl.utils import column_index_from_string
from src.services.log_service import registrar_log

from src.configs.configuration import (
    ORIGEM_DIR, BD_HIST_CFG
)
from src.models.database.repositories.tarifas_repository import TarifasRepository

# Utilitários de parsing
def _to_date(v) -> Optional[date]:
    if v is None:
        return None
    if isinstance(v, date) and not isinstance(v, datetime):
        return v
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, str):
        s = v.strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(s, fmt).date()
            except ValueError:
                pass
    return None

def _to_decimal(v) -> Optional[Decimal]:
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return None
    if isinstance(v, (int, float)):
        try:
            return Decimal(str(v))
        except InvalidOperation:
            return None
    if isinstance(v, Decimal):
        return v
    if isinstance(v, str):
        s = v.strip().replace(" ", "").replace(",", ".")
        try:
            return Decimal(s)
        except InvalidOperation:
            return None
    return None

def _to_text(v) -> Optional[str]:
    if v is None:
        return None
    s = str(v).strip()
    return s if s != "" else None

def _col(ws, row_idx: int, col_letter: str):
    col_idx = column_index_from_string(col_letter)
    return ws.cell(row=row_idx, column=col_idx).value

def _find_sheet(wb, contains: str) -> Optional[str]:
    tgt = contains.lower()
    for nm in wb.sheetnames:
        if tgt in nm.lower():
            return nm
    return None

# Mapeamento da aba "Int. Call Disc"
def _map_int_call_disc(ws_int, start_row=12):
    idd_out = {}
    idd_in = {}
    for r in range(start_row, ws_int.max_row + 1):
        plmn_out = _col(ws_int, r, "A")
        if plmn_out not in (None, ""):
            rate_out = _to_decimal(_col(ws_int, r, "G"))
            incr_out = _to_text(_col(ws_int, r, "H"))
            idd_out[str(plmn_out).strip()] = (rate_out, incr_out)
        plmn_in = _col(ws_int, r, "O")
        if plmn_in not in (None, ""):
            rate_in = _to_decimal(_col(ws_int, r, "U"))
            incr_in = _to_text(_col(ws_int, r, "V"))
            idd_in[str(plmn_in).strip()] = (rate_in, incr_in)
    return idd_out, idd_in

def _extract_rows_for_mysql_split(
    wb,
    base_contains="Base Disc Rate",
    int_contains="Int. Call Disc",
) -> Tuple[List[Tuple], List[Tuple]]:
    try:
        registrar_log("Iniciado", "Sucesso", f"Contagem de linhas a serem inseridas na Base de Dados.")
        base_name = _find_sheet(wb, base_contains)
        int_name = _find_sheet(wb, int_contains)
        if not base_name:
            raise RuntimeError("Aba 'Base Disc Rates' não encontrada.")
        if not int_name:
            raise RuntimeError("Aba 'Int. Call Disc' não encontrada.")
        
        ws_base = wb[base_name]
        ws_int = wb[int_name]
        
        idd_out_by_plmn, idd_in_by_plmn = _map_int_call_disc(ws_int, start_row=12)

        # 1) Mapeia IDD-OUT e IDD-IN separadamente
        registros_out: List[Tuple] = []
        registros_in: List[Tuple] = []
        for r in range(16, ws_base.max_row + 1):
            # TAP-OUT (Base Disc Rates)
            plmn_code_out = _col(ws_base, r, "A")
            if plmn_code_out not in (None, ""):
                registros_out.append((
                    str(plmn_code_out).strip(),
                    _to_text(_col(ws_base, r, "B")),
                    _to_text(_col(ws_base, r, "C")),
                    _to_date(_col(ws_base, r, "D")),
                    _to_date(_col(ws_base, r, "E")),
                    _to_decimal(_col(ws_base, r, "F")),
                    _col(ws_base, r, "G"),
                    _to_decimal(_col(ws_base, r, "M")),
                    _col(ws_base, r, "N"),
                    _to_decimal(_col(ws_base, r, "S")),
                    _col(ws_base, r, "T"),
                    _to_decimal(_col(ws_base, r, "Y")),
                    _to_decimal(_col(ws_base, r, "AA")),
                    _col(ws_base, r, "AB"),
                    *idd_out_by_plmn.get(str(plmn_code_out).strip(), (None, None)),
                ))

            # TAP-IN (Base Disc Rates)
            plmn_code_in = _col(ws_base, r, "AD")
            if plmn_code_in not in (None, ""):
                registros_in.append((
                    str(plmn_code_in).strip(),
                    _to_text(_col(ws_base, r, "AE")),
                    _to_text(_col(ws_base, r, "AF")),
                    _to_date(_col(ws_base, r, "AG")),
                    _to_date(_col(ws_base, r, "AH")),
                    _to_decimal(_col(ws_base, r, "AI")),
                    _col(ws_base, r, "AJ"),
                    _to_decimal(_col(ws_base, r, "AP")),
                    _col(ws_base, r, "AQ"),
                    _to_decimal(_col(ws_base, r, "AV")),
                    _col(ws_base, r, "AW"),
                    _to_decimal(_col(ws_base, r, "BB")),
                    _to_decimal(_col(ws_base, r, "BD")),
                    _col(ws_base, r, "BE"),
                    *idd_in_by_plmn.get(str(plmn_code_in).strip(), (None, None)),
                ))
        registrar_log("Finalizado", "Sucesso", f"Contagem de linhas a serem inseridas concluída. TAP-OUT: {len(registros_out)}, TAP-IN: {len(registros_in)} ")
        return registros_out, registros_in
    except Exception as e:
        registrar_log("Iniciado", "Erro", f"Falha na contagem de linhas: {e}")

# API do service (assinatura compatível) - Exportação para MySQL (duas tabelas)
def exportar_hist_tarifas_mysql_split(
    consolidado_path: Optional[str] = None,
    mysql_cfg: Dict[str, object] = BD_HIST_CFG,
    tabela_out: str = "hist_tarifas_roaming_int_tap_out",
    tabela_in: str = "hist_tarifas_roaming_int_tap_in",
    on_duplicate_update: bool = True,
    chunk_size: int = 1000,
) -> Tuple[int, int]:
    """
    Exporta os dados em duas tabelas (OUT e IN). Retorna (linhas_out, linhas_in).
    """
    repo = TarifasRepository(mysql_cfg)
    repo.connect()
    try:
        registrar_log("Iniciado", "Sucesso", f"Exportação de dados da planilha consolidada para a Base de Dados.")
        if consolidado_path:
            arquivos = [consolidado_path]
        else:
            if not os.path.exists(ORIGEM_DIR):
                raise RuntimeError(f"Diretório não encontrado: {ORIGEM_DIR}")
            arquivos = [
                os.path.join(ORIGEM_DIR, f)
                for f in os.listdir(ORIGEM_DIR)
                if f.lower().endswith(".xlsx")
            ]

        campos_out = (
            "plmn_code_tap_out, network_operator_tap_out, currency_tap_out, "
            "effective_date_tap_out, expiration_date_tap_out, "
            "moc_national_rate_tap_out, moc_national_increment_tap_out, "
            "mtc_rate_tap_out, mtc_increment_tap_out, "
            "moc_discount_home_rate_tap_out, moc_discount_home_increment_tap_out, "
            "sms_tap_out, gprs_rate_tap_out, gprs_increment_tap_out, "
            "idd_rate_tap_out, idd_increment_tap_out"
        )
        placeholders_out = ", ".join(["%s"] * 16)
        insert_sql_out = f"INSERT INTO {tabela_out} ({campos_out}) VALUES ({placeholders_out})"
        if on_duplicate_update:
            update_assignments_out = ", ".join(
                [f"{col} = VALUES({col})" for col in campos_out.replace(" ", "").split(",")]
            )
            insert_sql_out += f"ON DUPLICATE KEY UPDATE {update_assignments_out}"

        campos_in = (
            "plmn_code_tap_in, network_operator_tap_in, currency_tap_in, "
            "effective_date_tap_in, expiration_date_tap_in, "
            "moc_national_rate_tap_in, moc_national_increment_tap_in, "
            "mtc_rate_tap_in, mtc_increment_tap_in, "
            "moc_discount_home_rate_tap_in, moc_discount_home_increment_tap_in, "
            "sms_tap_in, gprs_rate_tap_in, gprs_increment_tap_in, "
            "idd_rate_tap_in, idd_increment_tap_in"
        )
        placeholders_in = ", ".join(["%s"] * 16)
        insert_sql_in = f"INSERT INTO {tabela_in} ({campos_in}) VALUES ({placeholders_in})"
        if on_duplicate_update:
            update_assignments_in = ", ".join(
                [f"{col} = VALUES({col})" for col in campos_in.replace(" ", "").split(",")]
            )
            insert_sql_in += f"ON DUPLICATE KEY UPDATE {update_assignments_in}"

        total_out = 0
        total_in = 0

        for arq in arquivos:
            wb = openpyxl.load_workbook(arq, data_only=True)
            rows_out, rows_in = _extract_rows_for_mysql_split(wb)

            if rows_out:
                for i in range(0, len(rows_out), chunk_size):
                    chunk = rows_out[i : i + chunk_size]
                    total_out += repo.insert_many(insert_sql_out, chunk)

            if rows_in:
                for i in range(0, len(rows_in), chunk_size):
                    chunk = rows_in[i : i + chunk_size]
                    total_in += repo.insert_many(insert_sql_in, chunk)

        return int(total_out), int(total_in)
    
    except Exception as e:
        registrar_log("Iniciado", "Erro", f"Erro na exportação de dados da planilha consolidada para a Base de Dados: {e}")

    finally:
        repo.close()