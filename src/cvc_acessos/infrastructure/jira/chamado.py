"""
Infra / Jira — preenchimento e envio do chamado no portal Service Desk.
Recebe sempre a 'page' do Jira (nao cria sessao/login).
"""

import re
import time

from cvc_acessos.infrastructure.config.config_app import (
    JIRA_TIPO_SOLICITUD, JIRA_COR_TITULO,
)


def _fechar_cookies(jira):
    for t in ["Apenas o necessário", "Aceitar tudo", "Only necessary",
              "Accept all"]:
        bb = jira.locator(f"button:has-text('{t}')")
        if bb.count() > 0:
            try:
                bb.first.click()
            except Exception:
                pass
            break


def preencher_formulario(jira, dados):
    """Preenche Resumen + Descripcion (cabecalho teal + corpo + link) + Tipo.
    NAO envia. Retorna True se preencheu."""
    _fechar_cookies(jira)
    if jira.locator("#summary").count() == 0:
        return False

    titulo, corpo, link = dados["titulo"], dados["corpo"], dados["link"]

    # Resumen
    jira.fill("#summary", titulo)

    # Descripcion
    ed = jira.locator("#ak-editor-textarea")
    ed.click()
    time.sleep(0.4)
    jira.keyboard.press("Control+a")
    jira.keyboard.press("Delete")
    time.sleep(0.3)
    if titulo:                       # titulo como CABECALHO
        jira.keyboard.type("## ")
        time.sleep(0.2)
        jira.keyboard.type(titulo)
        jira.keyboard.press("Enter")
    if corpo:
        jira.keyboard.type(corpo)
        jira.keyboard.press("Enter")
    jira.keyboard.type("Exibir resultados")
    jira.keyboard.press("Shift+Home")
    time.sleep(0.3)
    jira.keyboard.press("Control+k")
    time.sleep(1.2)
    if link:
        jira.keyboard.type(link)
        time.sleep(0.4)
        jira.keyboard.press("Enter")

    # cor teal no titulo
    jira.keyboard.press("Escape")
    time.sleep(0.3)
    try:
        jira.keyboard.press("Control+Home")
        jira.keyboard.press("Shift+End")
        time.sleep(0.3)
        jira.locator("[aria-label*='Cor do texto' i]").first.click()
        time.sleep(0.8)
        sw = jira.locator(f"[aria-label='{JIRA_COR_TITULO}']")
        if sw.count() > 0:
            sw.first.click()
        jira.keyboard.press("Escape")
        time.sleep(0.3)
    except Exception:
        pass

    # Tipo de solicitud
    tipo = jira.locator("#pf-undefined-cd-26")
    if tipo.count():
        try:
            tipo.click()
            time.sleep(0.6)
            jira.keyboard.type(JIRA_TIPO_SOLICITUD)
            time.sleep(1.6)
            ops = jira.locator("[role='option']")
            for i in range(ops.count()):
                if "form" in (ops.nth(i).inner_text() or "").lower():
                    ops.nth(i).click()
                    break
        except Exception:
            pass
    return True


def _status_chamado(jira):
    """Le o badge de status do chamado (texto curto). '?' se nao achar."""
    for sel in ["[data-test-id*='status' i]", "[class*='status' i] span"]:
        loc = jira.locator(sel)
        for i in range(min(loc.count(), 6)):
            t = (loc.nth(i).inner_text() or "").strip()
            if t and len(t) < 40:
                return t
    return "?"


def cancelar_chamado(jira, codigo, portal_base, transicao="Cancelado pelo Solicitante",
                     log=print):
    """CANCELA o chamado 'codigo' pelo portal (transicao do solicitante).
    Usado no modo DEMO p/ desfazer. VALIDADO ao vivo (GAAR-11..22):
      1. abre o chamado (portal_base + codigo)
      2. clica a transicao p/ ABRIR o popup — clique INSTAVEL, com RETRY (5x)
         ate o popup aparecer (placeholder 'Comentário opcional' OU 2+ botoes
         de mesmo rotulo)
      3. CONFIRMA no ULTIMO botao visivel de rotulo=transicao (o overlay
         renderiza por ultimo; o botao 'Cancelar' do popup NAO cancela o chamado)
      4. verifica pelo BADGE de status (reload). Retorna True se cancelou.
    """
    jira.goto(portal_base + codigo, wait_until="domcontentloaded")
    time.sleep(6)
    if jira.get_by_role("button", name=transicao, exact=False).count() == 0:
        return "cancel" in _status_chamado(jira).lower()
    aberto = False
    for _ in range(5):
        try:
            jira.get_by_role("button", name=transicao, exact=False).first.click()
        except Exception:
            pass
        for _ in range(8):
            if (jira.get_by_placeholder("Comentário opcional").count() > 0
                    or jira.get_by_role("button", name=transicao,
                                        exact=False).count() >= 2):
                aberto = True
                break
            time.sleep(0.5)
        if aberto:
            break
        time.sleep(1)
    cands = jira.get_by_role("button", name=transicao, exact=False)
    for i in reversed(range(cands.count())):
        try:
            if cands.nth(i).is_visible():
                cands.nth(i).click()
                break
        except Exception:
            pass
    time.sleep(4)
    jira.reload(wait_until="domcontentloaded")
    time.sleep(4)
    ok = "cancel" in _status_chamado(jira).lower()
    log(f"   Cancelamento {codigo}: {'OK' if ok else 'nao confirmado'}")
    return ok


def enviar_e_capturar_codigo(jira):
    """Clica em Enviar e captura o codigo do chamado (VALIDADO ao vivo: GAAR-x).
    Captura pela URL do chamado criado (.../portal/<n>/<CHAVE>?created=true);
    cai no corpo da pagina como fallback."""
    url_antes = jira.url
    jira.locator(
        "button:has-text('Enviar'), button:has-text('Crear'), "
        "button:has-text('Criar'), button:has-text('Send'), button[type='submit']"
    ).first.click()
    time.sleep(5)
    m = re.search(r"/portal/\d+/([A-Z]{2,8}-\d+)", jira.url)
    if m and jira.url != url_antes:
        return m.group(1)
    txt = jira.locator("body").inner_text()
    m2 = re.search(r"\b[A-Z]{2,8}-\d+\b", txt)
    return m2.group(0) if m2 else "??-?"
