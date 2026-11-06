import os
import shutil
from src.services.log_service import registrar_log
from src.utils.file_utils import ensure_dirs, limpar_diretorio
from src.configs.configuration import (
    ORIGEM_DIR, DESTINO_DIR, MODELO_DIR, SAIDA_DIR, PULSEFORCE_FILE,
    BD_HIST_CFG, REDE_PROCESSADOS_DIR, REDE_COMFONE
)
# Etapa 1 – Rede/Planilha/Consolidação
from src.services.rede_service import copiar_arquivos_validos
from src.services.planilha_service import carregar_pulseforce
from src.services.consolidacao_service import criar_planilha_consolidada, processar_planilhas
# Etapa 2 – Exportação para a Base de Dados
from src.services.mysql_export_service import exportar_hist_tarifas_mysql_split
# Etapa 3 – Export do Loader + E-mail de cadastro
from src.services.loader_export_service import executar_export_loader
from src.services.email_cadastro_operadoras_service import enviar_email_cadastro_operadoras
# Etapa 4 – Preenchimento do loader (XML) com regras e injeções
from src.services.preenchimento_loader_service import preencher_loader_multiplas_abas
# Etapa 5 - Subir loader atualizado no portal 
from src.services.load_loader_service import start_load_loader
# Etapa 6 - Envio do e-mail de notificação de atualização de tarifas
from src.services.email_final_service import enviar_email_notificacao_final

def executar_fluxo_principal() -> None:
    # ETAPA Complementar 1:
    # Garante diretórios locais
    try:
        ensure_dirs(ORIGEM_DIR, DESTINO_DIR, MODELO_DIR, SAIDA_DIR)
    except Exception as e: 
        registrar_log("Iniciado", "Erro", f"Algum dos diretórios essenciais do processamento não foi encontrado:{e}")

    # Limpeza inicial dos diretórios
    # Limpeza do diretório 'data/arquivos_origem'
    limpar_diretorio(ORIGEM_DIR)
    # Limpeza do diretório 'data/arquivos_destino'
    limpar_diretorio(DESTINO_DIR)
    # Limpeza do diretório 'data/saida'
    limpar_diretorio(SAIDA_DIR)
    # Limpar diretório de arquivos processados na REDE 
    # (os arquivos do mês anterior são mantidos até o próximo processamento)
    limpar_diretorio(REDE_PROCESSADOS_DIR)

    # ETAPA 1: Copiar arquivos válidos da rede, criar consolidada e processar planilhas
    try:
        copiar_arquivos_validos()  # Copiar/mover arquivos válidos da rede
        consolidado_path = criar_planilha_consolidada()   # Criar consolidada
        # Carregar Pulseforce e processar planilhas 
        if consolidado_path and os.path.exists(PULSEFORCE_FILE):
            tap_to_alias = carregar_pulseforce(PULSEFORCE_FILE)
            processar_planilhas(tap_to_alias)
            registrar_log("Finalizado", "Sucesso", f"ETAPA 1 concluída. Consolidação da planilha realizada.")
        else:
            registrar_log("Iniciado", "Erro", "Arquivo Pulseforce não encontrado. Encerrando.")
            return
    except Exception as e:
        registrar_log("Iniciado", "Erro", f"Falha na ETAPA 1 (Consolidação): {e}")
        return

    # ETAPA 2: Exportar histórico para Base de Dados e organiza arquivos
    try:
        # Exportação 
        linhas_out, linhas_in = exportar_hist_tarifas_mysql_split(
            consolidado_path,
            BD_HIST_CFG 
        )
        registrar_log(
            "Finalizado", "Sucesso",
            f"ETAPA 2 concluída. Exportação para Base de Dados. Linhas inseridas - TAP-OUT: {linhas_out} | TAP-IN: {linhas_in}"
        )
    except Exception as e:
        import traceback
        registrar_log("Finalizado", "Erro", f"Falha na ETAPA 2 (MySQL): {e}\nTRACE:\n{traceback.format_exc()}")

    # ETAPA Complementar 2:
    #Limpar diretório 'arquivos_origem' após processar planilha consolidada
    try:
        limpar_diretorio(ORIGEM_DIR)
        registrar_log("Iniciado", "Sucesso", "Limpeza do diretório 'arquivos_origem' após consolidação.")
    except Exception as e:
        registrar_log("Iniciado", "Erro", f"Erro na limpeza do diretório 'arquivos_origem' após consolidação: {e}")

    # Mover a consolidada para 'arquivos_origem' para as etapas seguintes
    try:
        if consolidado_path and os.path.exists(consolidado_path):
            novo_destino = os.path.join(ORIGEM_DIR, os.path.basename(consolidado_path))
            shutil.move(consolidado_path, novo_destino)
            registrar_log("Iniciado", "Sucesso", f"Arquivo consolidado movido para: {novo_destino}")
        else:
            registrar_log("Iniciado", "Erro", "Nenhum arquivo consolidado encontrado para mover.")
    except Exception as e:
        registrar_log("Iniciado", "Erro", f"Falha ao mover consolidado para 'arquivos_origem': {e}")

    # Limpar diretório de arquivo Comfone
    limpar_diretorio(REDE_COMFONE)
    registrar_log("Finalizado", "Sucesso", f"Diretório de arquivo Comfone limpo.")

    # ETAPA 3: Export do Loader + e-mail de cadastro (apenas se houver PLMNs rejeitados)
    try:
        resumo = executar_export_loader()  # {'path': Path, 'qtd': int, 'vazio': bool}
        if not resumo.get("vazio", True):
            used_file = enviar_email_cadastro_operadoras()
            registrar_log("Finalizado","Sucesso", f"E-mail de cadastro enviado. TXT anexado: {used_file}")
        else:
            registrar_log("Finalizado", "Sucesso", "Nenhum PLMN rejeitado. E-mail de cadastro NÃO será enviado.")
        registrar_log("Finalizado", "Sucesso", f"ETAPA 3 concluída. Export do loader e envio de e-mail de cadastro (se necessário).")
    except Exception as e:
        registrar_log("Finalizado", "Erro", f"Falha na ETAPA 3 (Export Loader/e-mail de cadastro): {e}")

    # ETAPA 4: Preencher (atualizar) o loader (XML)
    try:
        res = preencher_loader_multiplas_abas()  # {'xlsx': Path, 'xml': Path}
        registrar_log("Finalizado", "Sucesso", f"ETAPA 4 concluída. XML atualizado em: {res['xml']}")
    except Exception as e:
        registrar_log("Finalizado", "Erro", f"Falha na ETAPA 4 (Preenchimento do Loader): {e}")
    
    # ETAPA 5: Carregar o XML atualizado no portal
    try:
        xml_usado = start_load_loader()
        registrar_log("Finalizado", "Sucesso", f"ETAPA 5 concluída. Start Load disparado usando: {xml_usado.name}")
    except Exception as e:
        registrar_log("Finalizado", "Erro", f"Falha na ETAPA 5 (Start Load): {e}")
   
    # ETAPA 6: E-mail final de notificação com arquivo TAP
    try:
        arquivo_usado = enviar_email_notificacao_final()
        registrar_log("Finalizado", "Sucesso", f"ETAPA 6 concluída. E-mail final enviado. Anexo: {arquivo_usado.name}")
    except Exception as e:
        registrar_log("Finalizado", "Erro", f"Falha na ETAPA 6 (E-mail final): {e}")