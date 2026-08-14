"""
Infra / Jira via API (JSM REST) — cria e cancela o chamado SEM navegador.

Usado quando ha usuario + token no config (JIRA_API_ATIVA). Substitui o
fluxo de formulario do `chamado.py`:

  - criar_chamado(dados)  -> (codigo, link)   (POST /rest/servicedeskapi/request)
  - cancelar_chamado(cod) -> bool             (POST .../request/{cod}/transition)

A resposta da criacao ja traz o issueKey (GAAR-xx) e o link, entao NAO ha
scraping, espera de redirect, tela-pequena nem fallback de Enter.

Parametros (config.xml, bloco <jira>): usuario, token, service_desk_id,
request_type_id, transicao_cancelar_id. IDs default = portal do Bruna (GAAR),
apurados na API: serviceDesk 1984 / requestType 7550 / transicao 4.
"""

import base64
import json
import urllib.error
import urllib.request

from cvc_trata_forms.infrastructure.config.config_app import (
    JIRA_API_USUARIO, JIRA_API_TOKEN, JIRA_SITE, JIRA_SERVICE_DESK_ID,
    JIRA_REQUEST_TYPE_ID, JIRA_TRANSICAO_CANCELAR_ID,
)

_TIMEOUT = 30


def _auth_header():
    cred = f"{JIRA_API_USUARIO}:{JIRA_API_TOKEN}".encode()
    return "Basic " + base64.b64encode(cred).decode()


def _api(path, method="GET", body=None):
    """Chama a REST API. Retorna (status, obj|texto)."""
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(
        JIRA_SITE + path, data=data, method=method,
        headers={"Authorization": _auth_header(), "Accept": "application/json",
                 "Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=_TIMEOUT) as r:
            txt = r.read().decode()
            return r.status, (json.loads(txt) if txt.strip() else {})
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode(errors="replace")[:500]
    except Exception as e:  # noqa: BLE001
        return -1, repr(e)


def _montar_descricao(dados):
    """Descricao do chamado: corpo do e-mail + link 'Exibir resultados'."""
    partes = []
    if dados.get("corpo"):
        partes.append(dados["corpo"])
    if dados.get("link"):
        partes.append(f"Exibir resultados: {dados['link']}")
    return "\n\n".join(partes) or (dados.get("titulo") or "")


def criar_chamado(dados, log=print):
    """Cria o chamado via API. Retorna (codigo, link).
    Em caso de falha retorna ('??-?', '') e LOGA (o caller aborta)."""
    payload = {
        "serviceDeskId": str(JIRA_SERVICE_DESK_ID),
        "requestTypeId": str(JIRA_REQUEST_TYPE_ID),
        "requestFieldValues": {
            "summary": dados.get("titulo") or "Solicitud",
            "description": _montar_descricao(dados),
        },
    }
    st, res = _api("/rest/servicedeskapi/request", "POST", payload)
    if st in (200, 201) and isinstance(res, dict):
        codigo = res.get("issueKey")
        link = (res.get("_links") or {}).get("web") or ""
        if codigo:
            log(f"   [jira-api] chamado criado: {codigo}  ({link})")
            return codigo, link
    log(f"   [jira-api][ERRO] criar chamado falhou (status={st}): {res}")
    return "??-?", ""


def cancelar_chamado(codigo, log=print):
    """Cancela o chamado 'codigo' via API (transicao do solicitante).
    Usado no DEMO p/ desfazer. Retorna True se cancelou."""
    body = {
        "id": str(JIRA_TRANSICAO_CANCELAR_ID),
        "additionalComment": {"body": "Cancelamento automatico (modo DEMO)."},
    }
    st, res = _api(f"/rest/servicedeskapi/request/{codigo}/transition", "POST", body)
    ok = st in (200, 204)
    log(f"   [jira-api] cancelamento {codigo}: "
        f"{'OK' if ok else f'FALHOU (status={st}): {res}'}")
    return ok
