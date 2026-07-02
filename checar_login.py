"""
Verificador de sessao: confirma se OUTLOOK e JIRA estao logados no Edge
de automacao (porta 9222), antes de rodar qualquer fluxo.

USO:
    python checar_login.py

Tambem e usado como TRAVA DE SEGURANCA pelo fluxo principal:
    from checar_login import garantir_sessoes
    browser, outlook, jira = garantir_sessoes()   # lanca erro se algo cair
"""

import sys

from playwright.sync_api import sync_playwright

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

PORTA_CDP = "http://127.0.0.1:9222"

# Hosts de tela de login (se a aba cair pra ca, NAO esta logada)
LOGIN_HINTS = ("login.microsoftonline.com", "login.live.com",
               "login.microsoft.com", "id.atlassian.com", "/login")


def _achar_abas(browser):
    """Localiza as abas do Outlook e do Jira entre as abertas."""
    outlook = jira = None
    for ctx in browser.contexts:
        for pg in ctx.pages:
            u = pg.url.lower()
            if ("outlook" in u or "office.com" in u) and outlook is None:
                outlook = pg
            elif "atlassian" in u and jira is None:
                jira = pg
    return outlook, jira


def _esta_logado(page, positivos):
    """Heuristica generica de login.

    NAO logado se: a URL for de tela de login, OU houver campo de senha visivel.
    Logado se: passar nos testes acima E achar ao menos 1 sinal positivo
    (um seletor que so existe quando a sessao esta ativa).
    """
    url = page.url.lower()
    if any(h in url for h in LOGIN_HINTS):
        return False, "redirecionado para tela de login"
    if page.locator("input[type='password']:visible").count() > 0:
        return False, "campo de senha presente (tela de login)"
    for sel in positivos:
        try:
            if page.locator(sel).count() > 0:
                return True, "sessao ativa"
        except Exception:
            pass
    return False, "nao achei sinal de sessao ativa (pagina ainda carregando?)"


def checar(browser=None):
    """Retorna dict com o estado das duas sessoes."""
    fechar = False
    if browser is None:
        p = sync_playwright().start()
        try:
            browser = p.chromium.connect_over_cdp(PORTA_CDP)
        except Exception:
            # Edge fechado/porta caiu -> NAO deixar a sessao Playwright orfa
            # (senao poluiria o thread e quebraria o proximo login).
            try:
                p.stop()
            except Exception:
                pass
            raise
        browser._pw_session = p  # type: ignore[attr-defined]
        fechar = False  # nao fechamos o Edge do usuario

    outlook, jira = _achar_abas(browser)

    estado = {"outlook": {}, "jira": {}, "browser": browser,
              "page_outlook": outlook, "page_jira": jira}

    if outlook is None:
        estado["outlook"] = {"ok": False, "msg": "aba do Outlook nao encontrada"}
    else:
        ok, msg = _esta_logado(
            outlook,
            ["div[role='treeitem']", "div[role='option']", "div[role='navigation']"],
        )
        estado["outlook"] = {"ok": ok, "msg": msg, "url": outlook.url}

    if jira is None:
        estado["jira"] = {"ok": False, "msg": "aba do Jira nao encontrada"}
    else:
        ok, msg = _esta_logado(
            jira,
            ["#summary", "[data-testid*='profile']",
             "[aria-label*='conta' i]", "[aria-label*='account' i]",
             "input#request-type-select"],
        )
        estado["jira"] = {"ok": ok, "msg": msg, "url": jira.url}

    return estado


def garantir_sessoes():
    """Trava de seguranca: so retorna se AS DUAS sessoes estiverem OK."""
    e = checar()
    problemas = []
    if not e["outlook"].get("ok"):
        problemas.append(f"Outlook: {e['outlook'].get('msg')}")
    if not e["jira"].get("ok"):
        problemas.append(f"Jira: {e['jira'].get('msg')}")
    if problemas:
        raise RuntimeError(
            "Sessao(oes) nao logada(s):\n  - " + "\n  - ".join(problemas)
            + "\n\nAbra/atualize a(s) aba(s) no Edge de automacao e faca login."
        )
    return e["browser"], e["page_outlook"], e["page_jira"]


def main():
    print(">> Verificando sessoes no Edge de automacao (porta 9222)...\n")
    try:
        e = checar()
    except Exception as ex:  # noqa: BLE001
        print("[ERRO] Nao conectei no Edge. Ele foi aberto com abrir_edge.ps1?")
        print("       Detalhe:", ex)
        raise SystemExit(1)

    def linha(nome, d):
        marca = "OK   " if d.get("ok") else "FALHA"
        print(f"  [{marca}] {nome:8} -> {d.get('msg')}")
        if d.get("url"):
            print(f"           {d['url'][:70]}")

    print("=" * 70)
    linha("OUTLOOK", e["outlook"])
    linha("JIRA", e["jira"])
    print("=" * 70)

    if e["outlook"].get("ok") and e["jira"].get("ok"):
        print("\n>> Tudo logado! Pronto para rodar o fluxo.")
    else:
        print("\n>> Faca login na(s) aba(s) marcada(s) como FALHA e rode de novo.")
        raise SystemExit(1)


if __name__ == "__main__":
    main()
