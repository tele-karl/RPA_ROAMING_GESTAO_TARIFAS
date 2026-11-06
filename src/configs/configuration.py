import os
from pathlib import Path
from dotenv import load_dotenv

# Carrega variáveis de ambiente do arquivo .env na raiz do projeto
load_dotenv()

## Diretórios
# Diretórios locais
RAIZ_PROJETO = Path(__file__).resolve().parents[2]
DATA_DIR = os.path.join(RAIZ_PROJETO, os.getenv("DATA_DIR"))
ORIGEM_DIR = os.path.join(DATA_DIR, os.getenv("ORIGEM_SUBDIR"))
DESTINO_DIR = os.path.join(DATA_DIR, os.getenv("DESTINO_SUBDIR"))
MODELO_DIR = os.path.join(DATA_DIR, os.getenv("MODELO_SUBDIR"))
MODELO_PATH = os.path.join(MODELO_DIR, os.getenv("MODELO_FILENAME"))
MODELO_XML = os.path.join(MODELO_DIR, os.getenv("MODELO_XML"))
SAIDA_DIR = os.path.join(DATA_DIR, os.getenv("SAIDA_SUBDIR"))
LOGS_DIR = os.path.join(RAIZ_PROJETO, os.getenv("LOGS_DIR"))

# Diretórios de rede que estão no .env
REDE_COMFONE = os.getenv("REDE_COMFONE")
REDE_NOVOS_DIR = os.getenv("REDE_NOVOS_DIR")
REDE_PROCESSADOS_DIR = os.getenv("REDE_PROCESSADOS_DIR")
REDE_EXCEPCIONAIS_DIR = os.getenv("REDE_EXCEPCIONAIS_DIR")

# Arquivo Comfone
PULSEFORCE_FILE = os.path.join(REDE_COMFONE, os.getenv("PULSEFORCE_FILE"))

## Credenciais Base de Dados
BD_HIST_CFG_HOST = os.getenv("BD_HIST_CFG_HOST")
BD_HIST_CFG_PORT = int(os.getenv("BD_HIST_CFG_PORT"))
BD_HIST_CFG_DB = os.getenv("BD_HIST_CFG_DB")
BD_HIST_CFG_USER = os.getenv("BD_HIST_CFG_USER")
BD_HIST_CFG_PASSWORD = os.environ.get("DB_ROAM_HIST_PASS")

BD_HIST_CFG = {
    "host": BD_HIST_CFG_HOST,
    "port": BD_HIST_CFG_PORT,
    "database": BD_HIST_CFG_DB,
    "user": BD_HIST_CFG_USER,
    "password": BD_HIST_CFG_PASSWORD,
}

## Credenciais Base de Logs
RPA_LOG_HOST=os.getenv("RPA_LOG_HOST")
RPA_LOG_PORT=int(os.getenv("RPA_LOG_PORT"))
RPA_LOG_DB=os.getenv("RPA_LOG_DB")
RPA_LOG_USER=os.getenv("RPA_LOG_USER")
RPA_LOG_PASSWORD=os.environ.get("DB_ROAM_LOG_PASS")

RPA_LOG_CFG = {
    "host": RPA_LOG_HOST,
    "port": RPA_LOG_PORT,
    "database": RPA_LOG_DB,
    "user": RPA_LOG_USER,
    "password": RPA_LOG_PASSWORD
}

## SMTP
EMAIL_SMTP_HOST = os.getenv("EMAIL_SMTP_HOST")
EMAIL_SMTP_PORT = int(os.getenv("EMAIL_SMTP_PORT"))
EMAIL_SMTP_TIMEOUT = int(os.getenv("EMAIL_SMTP_TIMEOUT"))

## E-mails: templates e remetente
# Notificação final
EMAIL_KEYWORD = os.getenv("EMAIL_KEYWORD")            # p/ localizar TAP
ASSUNTO_EMAIL_NOT = os.getenv("ASSUNTO_EMAIL_NOT")    # Assunto determinado pelo nome do arquivo TAP "New Rates {period}"
TEXTO_EMAIL_NOT = os.getenv("TEXTO_EMAIL_NOT")        # corpo com {effective_date}

# Cadastro de operadoras
CADASTRO_KEYWORD = os.getenv("CADASTRO_KEYWORD")     # "plmns_rejeitados" (nome do TXT)
ASSUNTO_EMAIL_CAD = os.getenv("ASSUNTO_EMAIL_CAD")
TEXTO_EMAIL_CAD = os.getenv("TEXTO_EMAIL_CAD")

# Remetente
SENDER_EMAIL = os.getenv("SENDER_EMAIL")

## Portal ICT (login)
USUARIO_ICT = os.getenv("USUARIO_ICT")
SENHA_ICT = os.environ.get("ICT_ROAM_PASS")

# EMAIL_TESTE = os.getenv("GRUPO_CAD_CADASTRO")
# print("EMAIL_TESTE =", EMAIL_TESTE)

# print("REDE_COMFONE =", REDE_COMFONE)
# print("PULSEFORCE_FILE =", PULSEFORCE_FILE)
# print("Exists? =", os.path.exists(PULSEFORCE_FILE))

# print(MODELO_PATH)
# print(MODELO_DIR)
# print("Exists? =", os.path.exists(MODELO_PATH))

# print("LOGS_DIR =", LOGS_DIR)
# print("Exists? =", os.path.exists(LOGS_DIR))