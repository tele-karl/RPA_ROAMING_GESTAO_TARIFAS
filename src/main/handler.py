from src.controllers.processamento_consolidado_controller import executar_fluxo_principal
from src.services.log_service import registrar_log
from src.utils.network_auth import autenticar_rede_lagoa

def main():
    try:
        registrar_log("Iniciado", "Sucesso", "Iniciando execução do RPA (Controller).")
        
        # Autenticar na rede antes de acessar arquivos UNC
        if not autenticar_rede_lagoa():
            registrar_log("Finalizado", "Erro", "Falha na autenticação de rede. RPA não pode continuar.")
            return
        
        executar_fluxo_principal()
        registrar_log("Finalizado", "Sucesso", "Execução do RPA finalizada!")
    except Exception as e:
        registrar_log("Finalizado", "Erro", f"Falha na execução do RPA: {e}")