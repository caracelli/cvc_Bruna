"""
Leitor do config.xml (URLs, credenciais e regras do fluxo).
Centraliza tudo que muda por cliente/maquina. Se o config.xml nao existir,
usa valores padrao (defaults) para nao quebrar.
"""

import os
import xml.etree.ElementTree as ET

from cvc_trata_forms.infrastructure.sistema.caminhos import RAIZ

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


def _txt_opcional(caminho, default=""):
    """Como _txt, mas distingue 'tag ausente' de 'tag vazia': a tag PRESENTE e
    vazia vale como escolha do usuario (texto vazio), nao cai no default.
    Usado em campos que fazem sentido desligar deixando em branco."""
    if _root is None:
        return default
    el = _root.find(caminho)
    if el is None:
        return default
    return (el.text or "").strip()


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
JIRA_TRANSICAO_CANCELAR = _txt("jira/transicao_cancelar", "Cancelado pelo Solicitante")
# prefixo do codigo do chamado (ex.: "GAAR"). Se preenchido, a captura do
# codigo pos-envio SO aceita tokens que comecam com ele -> blinda contra pegar
# um "ABC-123" que estava no corpo do e-mail. Vazio = aceita qualquer prefixo.
JIRA_PREFIXO_CHAMADO = _txt("jira/prefixo_chamado", "")
# base do portal (…/portal/1984/) derivada da URL de criacao, p/ abrir o chamado
JIRA_PORTAL_BASE = URL_JIRA.split("/group/")[0].rstrip("/") + "/"

# ---- Jira via API (REST) ----
# Se houver usuario + token, o app cria/cancela o chamado pela JSM REST API
# (sem abrir a aba do Jira nem login SSO). Sem token, cai no fluxo de navegador.
JIRA_API_USUARIO = _txt("jira/usuario")
JIRA_API_TOKEN = _txt("jira/token")
# site Atlassian (scheme+host) derivado da URL do chamado
JIRA_SITE = ("/".join(URL_JIRA.split("/")[:3]) if "://" in URL_JIRA
             else "https://cvccorp.atlassian.net")
# IDs do portal (estaveis; apurados na API). Default = portal do Bruna (GAAR).
JIRA_SERVICE_DESK_ID = _txt("jira/service_desk_id", "1984")
JIRA_REQUEST_TYPE_ID = _txt("jira/request_type_id", "7550")
# id da transicao de cancelamento do solicitante (4 = "Cancelado pelo Solicitante")
JIRA_TRANSICAO_CANCELAR_ID = _txt("jira/transicao_cancelar_id", "4")
# liga a API so quando ha credencial (usuario + token)
JIRA_API_ATIVA = bool(JIRA_API_USUARIO and JIRA_API_TOKEN)

# ---- Fluxo ----
REMETENTE_FILTRO = _txt("fluxo/remetente_filtro", "Microsoft Forms")
SUBPASTA_DESTINO = _txt("fluxo/subpasta_destino", "Finalizados")
SO_NAO_LIDOS = _bool("fluxo/so_nao_lidos", True)
MARCAR_COMO_LIDO = _bool("fluxo/marcar_como_lido", True)
ENCERRAR_SESSOES_NO_FIM = _bool("fluxo/encerrar_sessoes_no_fim", True)
DRY_RUN = _bool("fluxo/dry_run", True)
# DEMO (self-cleaning): faz o ciclo REAL (cria chamado + encaminha + move) e
# DEPOIS DESFAZ (cancela o chamado + volta o e-mail ao Inbox nao lido). Serve
# pra DEMONSTRAR o fluxo sem deixar residuo. Tem prioridade sobre o DRY_RUN.
DEMO = _bool("fluxo/demo", False)
try:
    DEMO_QTD = int(_txt("fluxo/demo_qtd", "1") or "1")
except ValueError:
    DEMO_QTD = 1
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
# tag PRESENTE e vazia = sem separador (o OWA ja poe o cabecalho De/Para/
# Assunto acima do e-mail encaminhado). Por _txt cair no default, a linha
# voltava mesmo com <separador></separador> - o oposto do que o config diz.
ENCAMINHAR_SEPARADOR = _txt_opcional(
    "encaminhamento/separador",
    "--- Mensagem original (Microsoft Forms) abaixo ---",
)
