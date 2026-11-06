from __future__ import annotations

import smtplib
import socket
import time
import mimetypes
from dataclasses import dataclass
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable, Optional, Union, List, Tuple, Dict
import os

from src.configs.configuration import (
    EMAIL_SMTP_HOST, EMAIL_SMTP_PORT, EMAIL_SMTP_TIMEOUT
)

from src.services.log_service import registrar_log

@dataclass(frozen=True)
class SMTPConfig:
    host: str = EMAIL_SMTP_HOST
    port: int = EMAIL_SMTP_PORT
    timeout: int = EMAIL_SMTP_TIMEOUT

def ensure_list(value: Union[str, Iterable[str], None]) -> List[str]:
    """ 
    Garantir que um valor que pode ser string ou lista acabe sempre virando uma lista de strings.
    """
    if value is None:
        return []
    if isinstance(value, str):
        v = value.strip()
        return [v] if v else []
    out = []
    for v in value:
        if v is None:
            continue
        s = str(v).strip()
        if s:
            out.append(s)
    return out

def attach_file(msg: EmailMessage, filepath: Union[str, Path]) -> None:
    """
    Anexa arquivo detectando mime-type automaticamente.
    Aceita str ou Path.
    Dá erro mais amigável em caso de caminho inválido.
    """
    path = Path(filepath).expanduser()
    
    # Se vier relativo, resolve a partir do diretório atual
    resolved = (Path.cwd() / path).resolve() if not path.is_absolute() else path.resolve()
    
    if not resolved.exists():
        raise FileNotFoundError(
            "Arquivo para anexo não encontrado.\n"
            f" - Informado: {filepath}\n"
            f" - Resolvido: {resolved}\n"
            f" - Diretório atual (cwd): {Path.cwd()}\n"
            "Dica: use um caminho absoluto válido ou relativo ao projeto."
        )
    
    ctype, encoding = mimetypes.guess_type(str(resolved))
    if ctype is None or encoding is not None:
        ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)
    with open(resolved, "rb") as f:
        data = f.read()

    msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=resolved.name)

def build_message(
    sender: str,
    to: Union[str, Iterable[str]],
    subject: str,
    body_text: Optional[str] = None,
    body_html: Optional[str] = None,
    cc: Optional[Union[str, Iterable[str]]] = None,
    bcc: Optional[Union[str, Iterable[str]]] = None,
    reply_to: Optional[str] = None,
    attachments: Optional[Iterable[Union[str, Path]]] = None,
    headers: Optional[Dict[str, str]] = None,
) -> Tuple[EmailMessage, List[str]]:
    to_list = ensure_list(to)
    cc_list = ensure_list(cc)
    bcc_list = ensure_list(bcc)

    if not to_list and not cc_list and not bcc_list:
        raise ValueError("Nenhum destinatário informado (to/cc/bcc).")

    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = subject
    
    if reply_to:
        msg["Reply-To"] = reply_to
    
    if headers:
        for k, v in headers.items():
            msg[k] = v

    if body_text and body_html:
        msg.set_content(body_text)
        msg.add_alternative(body_html, subtype="html")
    elif body_html:
        msg.set_content("Este e-mail contém conteúdo em HTML. Abra em um cliente compatível.")
        msg.add_alternative(body_html, subtype="html")
    else:
        msg.set_content(body_text or "")

    if attachments:
        for path in attachments:
            attach_file(msg, path)

    recipients = to_list + cc_list + bcc_list
    return msg, recipients

def enviar_email(
    sender: str,
    to: Union[str, Iterable[str]],
    subject: str,
    body_text: Optional[str] = None,
    body_html: Optional[str] = None,
    cc: Optional[Union[str, Iterable[str]]] = None,
    bcc: Optional[Union[str, Iterable[str]]] = None,
    reply_to: Optional[str] = None,
    attachments: Optional[Iterable[Union[str, Path]]] = None,
    headers: Optional[Dict[str, str]] = None,
    max_retries: int = 3,
    backoff_seconds: float = 2.0,
    debug_smtp: bool = False,
    config: SMTPConfig | None = None,
) -> None:
    """Envio com retry e backoff."""
    if config is None:
        config = SMTPConfig()

    msg, recipients = build_message(
        sender=sender, to=to, subject=subject,
        body_text=body_text, body_html=body_html,
        cc=cc, bcc=bcc, reply_to=reply_to,
        attachments=attachments, headers=headers,
    )

    last_error: Exception | None = None

    for attempt in range(1, max_retries + 1):
        try:
            registrar_log("Iniciado", "Sucesso", f"Conectando em {config.host}:{config.port} (tentativa {attempt}/{max_retries})")
            with smtplib.SMTP(config.host, config.port, timeout=config.timeout) as server:
                if debug_smtp:
                    server.set_debuglevel(1)

                server.ehlo()
                server.send_message(msg, from_addr=sender, to_addrs=recipients)
            return
        
        except (smtplib.SMTPException, socket.error, OSError) as e:
            last_error = e
            registrar_log("Finalizado", "Erro", f"Falha ao enviar e-mail: {e!r}")
            if attempt < max_retries:
                registrar_log("Iniciado", "Sucesso", f"Aguarda {backoff_seconds}s para nova tentativa.")
                time.sleep(backoff_seconds)

    registrar_log("Finalizado", "Erro", f"Falha ao enviar e-mail após {max_retries} tentativas. Último erro: {last_error!r}")          
    raise

def carregar_grupos_do_env(prefixos) -> dict:
    """
    Lê variáveis do .env para todos os prefixos informados.
    prefixos pode ser string única ou lista de strings 
    """
    if isinstance(prefixos, str):
        prefixos = [prefixos]
    grupos = {}
    for key, value in os.environ.items():
        for prefixo in prefixos:
            if key.startswith(prefixo):
                # Pega última parte do grupo para definir nome
                partes = key.split("_")
                nome_grupo = partes[-1].strip()
                emails = [e.strip() for e in value.split(",") if e.strip()]
                if emails:
                    grupos[nome_grupo] = emails
    return grupos

def enviar_email_por_grupos(
    sender: str,
    subject: str,
    body_text: str,
    body_html: str,
    cc=None,
    attachments=None,
    headers=None,
    max_retries=3,
    backoff_seconds=2.0,
    debug_smtp=False,
    prefixos_grupos=None,
) -> None:
    """
    prefixos_grupos pode ser:
      - "GROUP_TAP_"    -> usa só esse prefixo
      - ["GROUP_TAP_", "GROUP_ROAMING"] -> usa múltiplos
    """
    
    grupos = carregar_grupos_do_env(prefixos_grupos)
    if not grupos:
        registrar_log("Finalizado", "Erro", f"Nenhum grupo encontrado para prefixos: {prefixos_grupos}")
        raise
    for nome, destinatarios in grupos.items():
        registrar_log(f"Iniciado", "Sucesso", f"Enviando para grupo '{nome}' -> {destinatarios}")
        enviar_email(
            sender=sender, to=destinatarios, subject=subject,
            body_text=body_text, body_html=body_html, cc=cc,
            attachments=attachments, headers=headers,
            max_retries=max_retries, backoff_seconds=backoff_seconds, debug_smtp=debug_smtp,
        )