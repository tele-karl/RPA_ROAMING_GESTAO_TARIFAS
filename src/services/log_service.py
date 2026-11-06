from src.configs.configuration import RPA_LOG_CFG
from src.models.database.repositories.log_repository import LogRepository

_repo = LogRepository(RPA_LOG_CFG)

RPA_NAME = "Roaming - Gestão de Tarifas - MVP1"

def registrar_log(processo: str, status: str, descricao: str):
    """
    processo: 'Iniciado' | 'Finalizado'
    status:   'Sucesso'  | 'Erro'
    descricao: texto livre (até 1000 chars)
    """
    if processo not in ("Iniciado", "Finalizado"):
        raise ValueError("processo deve ser 'Iniciado' ou 'Finalizado'.")
    if status not in ("Sucesso", "Erro"):
        raise ValueError("status deve ser 'Sucesso' ou 'Erro'.")
    try:
        _repo.inserir_log(RPA_NAME, processo, status, descricao)
    except Exception as e:
        # Print de fallback para troubleshooting pontual.
        print(f"[FALHA LOG DB] {e} | {processo=} {status=} {descricao=}")