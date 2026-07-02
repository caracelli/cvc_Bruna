"""
=============================================================================
 BOOTSTRAP DE SESSOES  -  CVC / Gestao de Acessos
=============================================================================
Prepara o ambiente ANTES de qualquer fluxo:

  1. Abre o Edge em modo depuracao (perfil dedicado) -- se ja estiver aberto,
     apenas reaproveita.
  2. Garante a aba do OUTLOOK e AGUARDA voce logar nela.
  3. Garante a aba do JIRA e AGUARDA voce logar nela.
  4. So termina quando AS DUAS estiverem logadas.

Nao cria chamado, nao move e-mail, nao altera nada. Apenas prepara o login.

 USO:
     python iniciar_sessoes.py
=============================================================================
"""

import os
import subprocess
import sys
import time
from urllib.request import urlopen

from playwright.sync_api import sync_playwright

# Reaproveita a deteccao de login ja validada
from checar_login import _esta_logado
from credenciais import obter_outlook
from login_outlook import login_outlook, _ja_logado, CredencialInvalida
from login_jira import login_jira, jira_logado
from winutil import trazer_edge_frente, esconder_edge

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# ===========================  CONFIG  =======================================
# Caminhos PORTAVEIS (detecta Edge + perfil isolado em %LOCALAPPDATA%)
from caminhos import EDGE, PERFIL, PORTA, achar_edge
# URLs vem do config.xml (configuravel por cliente)
from config_app import URL_OUTLOOK, URL_JIRA, INVISIVEL

# sinais positivos de sessao ativa (por portal)
POS_OUTLOOK = ["div[role='treeitem']", "div[role='option']",
               "div[role='navigation']"]
POS_JIRA = ["#summary", "input#request-type-select",
            "[data-testid*='profile']", "[aria-label*='conta' i]",
            "[aria-label*='account' i]"]

TIMEOUT_LOGIN_S = 300       # tempo maximo de espera por login (cada portal)
INTERVALO_S = 2             # frequencia de checagem
# ============================================================================


def _porta_ativa():
    try:
        urlopen(f"http://127.0.0.1:{PORTA}/json/version", timeout=2)
        return True
    except Exception:
        return False


def _desativar_salvar_senha():
    """Semeia o Preferences do perfil para o Edge NAO oferecer salvar senha
    (evita o popup 'Salvar sua senha'). So escreve se ainda nao existir, para
    nao sobrescrever um perfil em uso. Como o logoff apaga 'Default', no
    proximo inicio o arquivo nao existe e e criado com a opcao desligada."""
    import json
    default_dir = os.path.join(PERFIL, "Default")
    prefs = os.path.join(default_dir, "Preferences")
    if os.path.exists(prefs):
        return
    try:
        os.makedirs(default_dir, exist_ok=True)
        conteudo = {
            "credentials_enable_service": False,
            "profile": {
                "password_manager_enabled": False,
                "password_manager_leak_detection": False,
            },
        }
        with open(prefs, "w", encoding="utf-8") as f:
            json.dump(conteudo, f)
    except Exception:
        pass


def abrir_edge_se_preciso(url_inicial):
    """Abre o Edge ja apontando para 'url_inicial' (o Outlook).

    Assim a 1a aba que sobe e a do Outlook -- o Jira so sera aberto depois,
    no Passo 2, quando o Outlook ja estiver logado.
    """
    if _porta_ativa():
        print(f">> Edge ja esta aberto em depuracao (porta {PORTA}). Reusando.")
        return

    # Resolve o caminho do Edge (com varredura como fallback) e valida.
    edge = EDGE or achar_edge()
    if not edge:
        raise RuntimeError(
            "Microsoft Edge nao encontrado nesta maquina (nem em locais "
            "conhecidos, PATH, Registro ou varredura). Instale o Edge."
        )

    print(">> Abrindo o Edge (perfil dedicado, modo depuracao)...")
    print(f"   Edge: {edge}")
    os.makedirs(PERFIL, exist_ok=True)
    _desativar_salvar_senha()   # evita o popup "Salvar sua senha" do Edge
    args = [
        edge,
        f"--remote-debugging-port={PORTA}",
        f"--user-data-dir={PERFIL}",
        "--no-first-run",
        "--no-default-browser-check",
        # ---- supressao de popups/avisos do Edge ----
        "--hide-crash-restore-bubble",        # "Restaurar paginas"
        "--disable-session-crashed-bubble",
        "--disable-features=Translate,TranslateUI,msEdgeTranslate,"
        "msHubApps,msEdgeShoppingAssistant,msImplicitSignin,"
        "msEdgeSyncPromotion,EdgeSyncPromotion,ShowSyncPromo,"
        "msSpotlight,msEdgeWelcomePage",      # traducao, sync, signin promo...
        "--disable-sync",                      # nao oferecer sincronizar conta
        "--disable-popup-blocking",
        # mantem a pagina ATIVA mesmo com a janela escondida/minimizada
        "--disable-backgrounding-occluded-windows",
        "--disable-renderer-backgrounding",
        "--disable-background-timer-throttling",
        "--start-maximized",                   # abre a janela visivel/grande
        url_inicial,
    ]
    # CREATE_BREAKAWAY_FROM_JOB: faz o Edge sobreviver quando o script termina
    # (sem DETACHED_PROCESS, para a janela aparecer em primeiro plano).
    creationflags = 0
    if os.name == "nt":
        CREATE_NEW_PROCESS_GROUP = 0x00000200
        CREATE_BREAKAWAY_FROM_JOB = 0x01000000
        creationflags = CREATE_NEW_PROCESS_GROUP | CREATE_BREAKAWAY_FROM_JOB
    def _lancar(exe):
        cmd = [exe] + args[1:]
        try:
            subprocess.Popen(cmd, creationflags=creationflags, close_fds=True)
        except FileNotFoundError:
            raise                       # caminho do Edge invalido
        except OSError:
            subprocess.Popen(cmd)       # job nao permitiu breakaway -> sem flags

    try:
        _lancar(edge)
    except FileNotFoundError:
        # caminho do Edge ficou invalido -> varre a maquina e tenta de novo
        print("   Caminho do Edge invalido; varrendo a maquina...")
        edge2 = achar_edge()
        if not edge2:
            raise RuntimeError("Microsoft Edge nao encontrado na maquina.")
        print(f"   Caminho do Edge ajustado: {edge2}")
        _lancar(edge2)

    for _ in range(30):
        if _porta_ativa():
            print("   Edge no ar.")
            return
        time.sleep(1)
    raise RuntimeError("O Edge nao subiu a porta de depuracao a tempo.")


def fechar_abas(browser, chaves):
    """Fecha quaisquer abas cuja URL contenha uma das 'chaves'.

    Usado para garantir que o Jira NAO esteja aberto durante o passo do
    Outlook (ex.: se o Edge restaurou a sessao anterior com a aba do Jira).
    """
    for ctx in browser.contexts:
        for pg in list(ctx.pages):
            if any(k in pg.url.lower() for k in chaves):
                try:
                    pg.close()
                except Exception:
                    pass


def _achar_aba(browser, chaves):
    for ctx in browser.contexts:
        for pg in ctx.pages:
            u = pg.url.lower()
            if any(k in u for k in chaves):
                return pg
    return None


def garantir_aba(browser, chaves, url):
    """Retorna a aba do portal; se nao existir, abre uma nova com a URL."""
    pg = _achar_aba(browser, chaves)
    if pg is None:
        ctx = browser.contexts[0] if browser.contexts else browser.new_context()
        pg = ctx.new_page()
        pg.goto(url, wait_until="domcontentloaded")
    return pg


def aguardar_login(page, nome, positivos):
    """Bloqueia ate a aba 'page' estar logada (ou estourar o timeout)."""
    ok, _ = _esta_logado(page, positivos)
    if ok:
        print(f"   [OK] {nome} ja esta logado.")
        return True

    print(f"   [..] Aguardando login no {nome} "
          f"-> FACA LOGIN na janela do Edge.")
    page.bring_to_front()
    fim = time.time() + TIMEOUT_LOGIN_S
    ult = ""
    while time.time() < fim:
        ok, msg = _esta_logado(page, positivos)
        if ok:
            print(f"   [OK] {nome} logado!")
            return True
        if msg != ult:           # so reimprime quando o status muda
            restante = int(fim - time.time())
            print(f"        ...{nome}: {msg} (restam ~{restante}s)")
            ult = msg
        time.sleep(INTERVALO_S)

    print(f"   [X] Tempo esgotado esperando login no {nome}.")
    return False


def _abortar_credencial(e, popup_proc=None):
    """Mensagem clara quando o e-mail/senha esta incorreto, e encerra o
    fechador de popups. Nao cai para o modo manual (seria espera inutil)."""
    print("\n" + "!" * 70)
    print(f"  [ERRO] {e}")
    print("  As credenciais salvas estao INCORRETAS.")
    print("  Corrija com:  python salvar_credenciais.py")
    print("  (ou edite o arquivo credenciais.xml) e rode novamente.")
    print("!" * 70)
    if popup_proc is not None:
        try:
            popup_proc.terminate()
        except Exception:
            pass


_PW = None   # sessao Playwright atual (fechada no fim por main())


def _main_impl():
    global _PW
    print("=" * 70)
    print("  BOOTSTRAP DE SESSOES  -  CVC / Gestao de Acessos")
    print("=" * 70)

    abrir_edge_se_preciso(URL_OUTLOOK)

    p = sync_playwright().start()
    _PW = p
    browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{PORTA}")

    time.sleep(1.5)
    if INVISIVEL:
        esconder_edge()        # roda invisivel; so aparece no MFA
    else:
        trazer_edge_frente()   # assim que abrir, traz SO o Edge pra frente

    # Fechador de popups do Edge em PROCESSO SEPARADO (UIA e lento/bloqueante;
    # rodando isolado nao trava o login). Encerrado no fim.
    popup_proc = None
    try:
        fp = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "fechar_popups.py")
        popup_proc = subprocess.Popen([sys.executable, fp, "--loop"])
    except Exception:
        popup_proc = None

    # ---- PASSO 1: OUTLOOK (abre e aguarda login ANTES de tocar no Jira) ----
    print("\n" + "-" * 70)
    print("  PASSO 1/2  -  OUTLOOK")
    print("-" * 70)
    # Garante que o Jira NAO esteja aberto ainda (se a sessao foi restaurada)
    fechar_abas(browser, ["atlassian"])
    outlook = garantir_aba(
        browser,
        ["outlook", "office.com", "login.microsoft", "login.live"],
        URL_OUTLOOK,
    )
    try:
        outlook.bring_to_front()
    except Exception:
        pass
    if INVISIVEL:
        esconder_edge()        # mantem escondido durante o login
    else:
        trazer_edge_frente()   # traz a janela do Edge pra frente (ver MFA)

    # DUAS FRENTES: com credencial salva -> login automatico (fallback manual);
    #               sem credencial -> espera manual.
    email, senha = obter_outlook()
    ok_outlook = False
    if email and senha and not _ja_logado(outlook):
        print(f"   Credencial encontrada para {email}.")
        print("   Tentando login automatico (o MFA sera com voce)...")
        try:
            ok_outlook = login_outlook(outlook, email, senha, log=print)
        except CredencialInvalida as e:
            _abortar_credencial(e, popup_proc)
            return
        except Exception as e:
            print(f"   (login automatico falhou: {e})")
            ok_outlook = False
        if not ok_outlook:
            print("   -> Caindo para modo MANUAL (faca login na janela do Edge).")
            ok_outlook = aguardar_login(outlook, "OUTLOOK", POS_OUTLOOK)
    else:
        if not (email and senha):
            print("   (sem credencial salva - modo MANUAL. "
                  "Dica: rode 'python salvar_credenciais.py' p/ automatizar)")
        ok_outlook = aguardar_login(outlook, "OUTLOOK", POS_OUTLOOK)

    # apos o MFA do Outlook, esconde de novo (Jira loga escondido tambem)
    if INVISIVEL and ok_outlook:
        esconder_edge()

    if not ok_outlook:
        print("\n>> Outlook nao logou. O Jira NAO sera aberto. "
              "Faca o login e rode de novo.")
        print("=" * 70)
        return

    # ---- PASSO 2: JIRA (so agora, com o Outlook ja logado) -----------------
    print("\n" + "-" * 70)
    print("  PASSO 2/2  -  JIRA  (Outlook ja logado, abrindo o Jira agora)")
    print("-" * 70)
    jira = garantir_aba(browser, ["atlassian"], URL_JIRA)

    # DUAS FRENTES tambem no Jira: automatico (SSO) com fallback manual.
    email, senha = obter_outlook()
    ok_jira = False
    if email and senha and not jira_logado(jira):
        print("   Tentando login automatico no Jira (SSO Microsoft)...")
        try:
            ok_jira = login_jira(jira, email, senha, log=print)
        except CredencialInvalida as e:
            _abortar_credencial(e, popup_proc)
            return
        except Exception as e:
            print(f"   (login Jira falhou: {e})")
            ok_jira = False
        if not ok_jira:
            print("   -> Caindo para modo MANUAL (faca login do Jira na janela).")
            ok_jira = aguardar_login(jira, "JIRA", POS_JIRA)
    else:
        ok_jira = aguardar_login(jira, "JIRA", POS_JIRA)

    # Encerra o processo fechador de popups.
    if popup_proc is not None:
        try:
            popup_proc.terminate()
        except Exception:
            pass

    print("\n" + "=" * 70)
    if ok_outlook and ok_jira:
        print("  PRONTO! Outlook e Jira logados. Ambiente preparado.")
        print("  Proximo passo: rodar o fluxo (em DRY-RUN) -> python fluxo.py")
    else:
        faltam = []
        if not ok_outlook:
            faltam.append("Outlook")
        if not ok_jira:
            faltam.append("Jira")
        print(f"  ATENCAO: ainda falta logar -> {', '.join(faltam)}")
        print("  Rode novamente apos efetuar o login.")
    print("=" * 70)

    # modo invisivel: apos logar, esconde o Edge de novo
    if INVISIVEL and ok_outlook and ok_jira:
        esconder_edge()


def main():
    """Faz o bootstrap e SEMPRE fecha a sessao Playwright no fim (sem fechar
    o Edge). Importante quando chamado em-processo (ex.: app de bandeja),
    para nao conflitar com a proxima sessao Playwright (ciclos do monitor)."""
    global _PW
    _PW = None
    try:
        _main_impl()
    finally:
        if _PW is not None:
            try:
                _PW.stop()
            except Exception:
                pass
            _PW = None


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:  # noqa: BLE001
        print("\n[ERRO] " + str(ex))
        raise SystemExit(1)
