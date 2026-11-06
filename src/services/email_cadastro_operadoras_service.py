import os
from pathlib import Path
from html import escape

from src.services.email_service import enviar_email_por_grupos
from src.services.log_service import registrar_log
from src.configs.configuration import (
    SAIDA_DIR, SENDER_EMAIL, CADASTRO_KEYWORD, ASSUNTO_EMAIL_CAD, TEXTO_EMAIL_CAD,
)

def encontrar_arquivo_rejeitados(saida_dir: Path, keyword: str) -> Path:
    """
    Procura o arquivo .txt que contenha 'plmns_rejeitados' no nome.
    Se houver mais de um, pega o mais recente (por data de modificação).
    """
    candidatos = list(saida_dir.glob(f"*{keyword}*.txt"))
    if not candidatos:
        registrar_log("Finalizado", "Erro", f"Nenhum arquivo '*{keyword}*.txt' encontrado em: {saida_dir.resolve()}")
        raise 
    candidatos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidatos[0]

def ler_texto_com_fallback(path: Path) -> str:
    """
    Lê o conteúdo do TXT com fallback de encoding.
    """
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="latin-1", errors="replace")

def enviar_email_cadastro_operadoras() -> Path:
    """
    Procura o TXT de 'plmns_rejeitados' em DATA/saida/ e envia e-mails de cadastro.
    Retorna o path do arquivo usado (para logging).
    """
    saida_dir = Path(SAIDA_DIR)
    saida_dir.mkdir(parents=True, exist_ok=True)

    arquivo_txt = encontrar_arquivo_rejeitados(saida_dir, CADASTRO_KEYWORD)

    conteudo = ler_texto_com_fallback(arquivo_txt).strip()
    body_text = f"{TEXTO_EMAIL_CAD}\n\n{conteudo}"

    conteudo_html = escape(conteudo).replace("\n", "<br/>")
    body_html = f"""
    <p>{escape(TEXTO_EMAIL_CAD)}</p>
    <p><strong>{conteudo_html}</strong></p>
    """

    # prefixos dos grupos (mantendo seu comportamento original)
    try:
        enviar_email_por_grupos(
            sender=SENDER_EMAIL,
            # to=to,
            cc=[c.strip() for c in os.getenv("GRUPO_ROAMING").split(',')],
            # bcc=bcc,
            subject=ASSUNTO_EMAIL_CAD,
            body_text=body_text,
            body_html=body_html,
            attachments=[arquivo_txt],
            prefixos_grupos="GRUPO_CAD_"
        )
        registrar_log("Finalizado", "Sucesso", f"E-mails para cadastro de operadoras enviado com sucesso!")
        return arquivo_txt
    except Exception as e:
        registrar_log("Finalizado", "Erro", f"Falha no envio de e-mails para cadastro de operadoras: {e}")    