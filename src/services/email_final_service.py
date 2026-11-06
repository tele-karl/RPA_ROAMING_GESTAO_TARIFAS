import re
import os
from pathlib import Path
from datetime import date

from src.configs.configuration import (
    EMAIL_KEYWORD, ASSUNTO_EMAIL_NOT, TEXTO_EMAIL_NOT,
    SENDER_EMAIL, ORIGEM_DIR
)
from src.services.log_service import registrar_log
from src.services.email_service import enviar_email_por_grupos

# captura "Jan-26", "Apr-26", etc.
PERIOD_REGEX = re.compile(r"\b([A-Z][a-z]{2})-(\d{2})\b")

# Mapeamento de abreviações de meses em inglês -> número
MONTH_MAP = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4,
    "May": 5, "Jun": 6, "Jul": 7, "Aug": 8,
    "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12
}

def _encontrar_arquivo_tap() -> Path:
    """
    Localiza o arquivo que contenha 'TAP' nome mais recente em data/arquivos_origem.
    """
    origem = Path(ORIGEM_DIR)
    candidatos = sorted(
        origem.glob(f"*{EMAIL_KEYWORD}*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )
    if not candidatos:
        registrar_log("Iniciado", "Erro", f"Nenhum arquivo '*{EMAIL_KEYWORD}*.xlsx' encontrado em {origem}.")
        raise 
    return candidatos[0]

def _extrair_periodo_mes_ano(nome_arquivo: str) -> tuple[str, int, int]:
    """
    Extrai o período no formato 'Mon-YY' do nome do arquivo.
    Retorna: (periodo_str, month_int, year_int)
      - periodo_str: 'Jan-26', 'Apr-26', ...
      - month_int: 1..12
      - year_int: 4 dígitos (ex.: 2026)
    """
    m = PERIOD_REGEX.search(nome_arquivo)
    if not m:
        registrar_log("Iniciado", "Erro", f"Não foi possível identificar o período 'Mon-YY' no nome do arquivo")
        raise 
    mon_abbr, yy = m.group(1), int(m.group(2))
    if mon_abbr not in MONTH_MAP:
        registrar_log("Iniciado", "Erro", f"Mês inválido: {mon_abbr}")
        raise 

    month = MONTH_MAP[mon_abbr]
    year = 2000 + yy    # assume século 2000 (ex.: '26' -> 2026)
    period_str = f"{mon_abbr}-{yy:02d}"
    return period_str, month, year

def _montar_assunto(period_str: str) -> str:
    """
    Monta o assunto usando o template do .env.
    Ex.: "New Rates {period}" -> "New Rates Jan-26"
    """
    return ASSUNTO_EMAIL_NOT.replace("{period}", period_str)

def _montar_body_text(effective_date: date) -> str:
    """Constroi o corpo do texto através da variável de ambiente padrão TEXTO_EMAIL"""
    eff_str = effective_date.strftime("%d/%m/%Y")
    texto = TEXTO_EMAIL_NOT.replace("{effective_date}", eff_str)
    # Converte '\n' literais para quebras reais
    return texto.replace("\\n", "\n")

def _converte_text_to_html(texto: str) -> str:
    """
    Converte o texto em HTML simples, respeitando parágrafos em linhas vazias.
    """
    # Divide em parágrafos por linhas em branco duplas
    paragrafos = [p.strip() for p in texto.split("\n\n") if p.strip()]
    html = "\n".join(
        f"<p>{p.replace('\n', '<br>')}</p>" for p in paragrafos
    )
    return html

def enviar_email_notificacao_final() -> Path:
    """
    Pipeline da notificação final:
      - Localiza TAP
      - Extrai período
      - Gera assunto/corpos
      - Anexa TAP
      - Envia aos grupos TAP + Roaming
    """
    arquivo_tap = _encontrar_arquivo_tap()

    period_str, mes, ano = _extrair_periodo_mes_ano(arquivo_tap.name)
    effective_date = date(ano, mes, 1)

    subject = _montar_assunto(period_str)
    body_text = _montar_body_text(effective_date)
    body_html = _converte_text_to_html(body_text)

    try:
        enviar_email_por_grupos(
            sender=SENDER_EMAIL,
            # to=to,
            cc=[c.strip() for c in os.getenv("GRUPO_ROAMING").split(',')],
            # bcc=bcc,
            subject=subject,
            body_text=body_text,
            body_html=body_html,
            attachments=[arquivo_tap],
            prefixos_grupos= ["GRUPO_TAP_", "GRUPO_CAD_"] 
        )
        registrar_log("Finalizado", "Sucesso", f"E-mails de notificação enviados com sucesso!")
        return arquivo_tap
    except Exception as e:
        registrar_log("Finalizado", "Erro", f"Falha no envio de e-mails de notificação: {e}") 