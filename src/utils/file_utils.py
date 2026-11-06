import os
import shutil
from typing import Iterable

from src.services.log_service import registrar_log

def limpar_diretorio(diretorio: str) -> None:
    """
    Remove todo o conteúdo do diretório (arquivos e subpastas)
    e registra logs de sucesso/erro da operação.
    """
    registrar_log("Iniciado", "Sucesso",
                  f"Iniciando limpeza do diretório: {diretorio}")

    if not os.path.exists(diretorio):
        os.makedirs(diretorio, exist_ok=True)
        return
    
    for nome in os.listdir(diretorio):
        caminho = os.path.join(diretorio, nome)
        try:
            if os.path.isfile(caminho) or os.path.islink(caminho):
                os.remove(caminho)
            elif os.path.isdir(caminho):
                shutil.rmtree(caminho)
        except Exception as e:
            print("[WARN]", f"Falha ao remover '{caminho}': {e}")

def ensure_dirs(*dirs: Iterable[str]) -> None:
    for d in dirs:
        os.makedirs(d, exist_ok=True)