"""
Encerra (LOGOFF REAL) as sessoes do Outlook e do Jira.

IMPORTANTE: so fechar o Edge NAO desloga — os cookies ficam salvos no perfil
e a proxima abertura entra direto. Logoff de verdade =
  1. visitar as URLs de logout (invalida a sessao no servidor), e
  2. LIMPAR os cookies do perfil (garante que a proxima vez peca login),
  3. e so entao fechar o Edge.

USO (standalone):
    python encerrar_sessoes.py
"""

import os
import subprocess
import time

from cvc_trata_forms.infrastructure.sistema.caminhos import PERFIL, MARCADOR_PERFIL

PORTA_CDP = "http://127.0.0.1:9222"
URL_LOGOUT_MS = "https://login.microsoftonline.com/common/oauth2/v2.0/logout"
URL_LOGOUT_ATLASSIAN = "https://id.atlassian.com/logout"


def _apagar_sessao_disco(log=print):
    """Apaga a pasta de DADOS do perfil (logoff GARANTIDO).

    So apagar cookies NAO basta: o token de sessao tambem fica em Local Storage
    / IndexedDB. Por isso removemos a pasta 'Default' inteira (cookies + tokens
    + storage). O Edge a recria vazia na proxima abertura -> sempre pede login.
    So funciona com o Edge JA FECHADO (senao os arquivos ficam travados).
    """
    import shutil
    alvo = os.path.join(PERFIL, "Default")
    if not os.path.exists(alvo):
        log("   (perfil ja estava limpo)")
        return
    for tentativa in range(5):
        try:
            shutil.rmtree(alvo)
            log("   Sessao em disco apagada (perfil 'Default' removido).")
            return
        except Exception:
            time.sleep(1)   # aguarda o Edge liberar os arquivos
    log("   [aviso] nao consegui remover toda a pasta do perfil (arquivos "
        "travados?). Garanta que o Edge esta fechado e tente de novo.")


def fechar_edge_automacao(log=print):
    """Fecha o Edge do perfil de automacao (encerra as abas)."""
    ps = (
        "$a=Get-CimInstance Win32_Process -Filter \"Name='msedge.exe'\" | "
        f"Where-Object {{ $_.CommandLine -like '*{MARCADOR_PERFIL}*' }}; "
        "if($a){ $a | ForEach-Object { "
        "Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } }"
    )
    subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                   capture_output=True, creationflags=0x08000000)  # CREATE_NO_WINDOW
    time.sleep(1)
    log("   Edge de automacao fechado.")


def logoff_completo(browser=None, log=print):
    """Logoff REAL: logout no servidor + limpa cookies + fecha o Edge."""
    pw = None
    if browser is None:
        try:
            from playwright.sync_api import sync_playwright
            pw = sync_playwright().start()
            browser = pw.chromium.connect_over_cdp(PORTA_CDP)
        except Exception:
            log("   (Edge nao estava aberto; garantindo logoff em disco)")
            if pw is not None:
                try:
                    pw.stop()      # NAO deixar a sessao Playwright orfa
                except Exception:
                    pass
            fechar_edge_automacao(log=log)
            _apagar_sessao_disco(log=log)
            return

    try:
        # 1) Visita as URLs de logout (invalida sessao no servidor)
        ctx = browser.contexts[0] if browser.contexts else None
        if ctx is not None:
            pg = ctx.pages[0] if ctx.pages else ctx.new_page()
            for url in (URL_LOGOUT_MS, URL_LOGOUT_ATLASSIAN):
                try:
                    pg.goto(url, wait_until="domcontentloaded", timeout=20000)
                    time.sleep(2)
                except Exception:
                    pass
            log("   Logout solicitado (Microsoft + Atlassian).")

        # 2) Limpa cookies de TODOS os contextos (forca login na proxima vez)
        for ctx in browser.contexts:
            try:
                ctx.clear_cookies()
            except Exception:
                pass
        log("   Cookies do perfil limpos.")
    except Exception as e:  # noqa: BLE001
        log(f"   (aviso durante o logout: {e})")
    finally:
        if pw is not None:
            try:
                pw.stop()
            except Exception:
                pass

    # 3) Fecha o Edge e 4) apaga cookies do disco (logoff garantido)
    fechar_edge_automacao(log=log)
    _apagar_sessao_disco(log=log)
    log("   Logoff concluido: a proxima abertura vai pedir login.")


if __name__ == "__main__":
    print(">> Logoff REAL (logout + limpar cookies + fechar Edge)...")
    logoff_completo()
    print(">> Pronto.")
