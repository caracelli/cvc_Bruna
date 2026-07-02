"""
Leitor do config.xml (URLs, credenciais e regras do fluxo).
Centraliza tudo que muda por cliente/maquina. Se o config.xml nao existir,
usa valores padrao (defaults) para nao quebrar.
"""

import os
import xml.etree.ElementTree as ET

from cvc_acessos.infrastructure.sistema.caminhos import RAIZ

ARQ = os.path.join(RAIZ, "config.xml")

_root = None
if os.path.exists(ARQ):
    try:
        _root = ET.parse(ARQ).getroot()
    except Exception:
        _root = None


def _txt(caminho, default=""):
    if _root is None:
        return default
    el = _root.find(caminho)
    if el is not None and el.text and el.text.strip():
        return el.text.strip()
    return default


def _bool(caminho, default=True):
    return _txt(caminho, str(default)).strip().lower() in ("true", "1", "sim", "yes")


# ---- Outlook / conta Microsoft ----
OUTLOOK_EMAIL = _txt("outlook/email")
OUTLOOK_SENHA = _txt("outlook/senha")
URL_OUTLOOK = _txt("outlook/url", "https://outlook.office.com/mail/")
CAIXA = _txt("outlook/caixa", "Gestão de Acessos")

# ---- Jira ----
URL_JIRA = _txt(
    "jira/url_chamado",
    "https://cvccorp.atlassian.net/servicedesk/customer/"
    "portal/1984/group/3223/create/7550",
)
JIRA_TIPO_SOLICITUD = _txt("jira/tipo_solicitud", "forms")
JIRA_COR_TITULO = _txt("jira/cor_titulo", "Verde-azulado forte")
JIRA_CAT_PREFIXO = _txt("jira/categoria_prefixo", "JIRA -")

# ---- Fluxo ----
REMETENTE_FILTRO = _txt("fluxo/remetente_filtro", "Microsoft Forms")
SUBPASTA_DESTINO = _txt("fluxo/subpasta_destino", "Finalizado")
SO_NAO_LIDOS = _bool("fluxo/so_nao_lidos", True)
MARCAR_COMO_LIDO = _bool("fluxo/marcar_como_lido", True)
ENCERRAR_SESSOES_NO_FIM = _bool("fluxo/encerrar_sessoes_no_fim", True)
DRY_RUN = _bool("fluxo/dry_run", True)
try:
    INTERVALO_MIN = int(_txt("fluxo/intervalo_min", "5") or "5")
except ValueError:
    INTERVALO_MIN = 5
INVISIVEL = _bool("fluxo/invisivel", False)

# ---- Encaminhamento (notifica o grupo com o nº do chamado) ----
ENCAMINHAR_ATIVO = _bool("encaminhamento/ativo", False)
_dest = _txt("encaminhamento/destinatarios", "")
ENCAMINHAR_DESTINATARIOS = [e.strip() for e in _dest.replace(";", ",").split(",")
                            if e.strip()]
ENCAMINHAR_ASSUNTO = _txt("encaminhamento/assunto_template", "{ticket} - {assunto}")
ENCAMINHAR_PREFIXO = _txt("encaminhamento/prefixo_corpo", "Chamado aberto: ")
ENCAMINHAR_SEPARADOR = _txt(
    "encaminhamento/separador",
    "--- Mensagem original (Microsoft Forms) abaixo ---",
)
