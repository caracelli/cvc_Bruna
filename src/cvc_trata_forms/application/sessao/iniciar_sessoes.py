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
from cvc_trata_forms.application.sessao.checar_login import _esta_logado
from cvc_trata_forms.infrastructure.credenciais.credenciais import obter_outlook
from cvc_trata_forms.infrastructure.outlook.login_outlook import (
    login_outlook, _ja_logado, CredencialInvalida,
)
from cvc_trata_forms.infrastructure.jira.login_jira import login_jira, jira_logado
from cvc_trata_forms.infrastructure.sistema.winutil import trazer_edge_frente, esconder_edge

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# ===========================  CONFIG  =======================================
# Caminhos PORTAVEIS (detecta Edge + perfil isolado em %LOCALAPPDATA%)
from cvc_trata_forms.infrastructure.sistema.caminhos import EDGE, PERFIL, PORTA, achar_edge
# URLs vem do config.xml (configuravel por cliente)
from cvc_trata_forms.infrastructure.config.config_app import (
    URL_OUTLOOK, URL_JIRA, INVISIVEL, JIRA_API_ATIVA,
)

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


def _log_layout(msg):
    """Loga direto no cvc_bandeja.log (o login roda fora do redirect de stdout,
    entao print() nao apareceria no log principal)."""
    try:
        import time as _t
        from cvc_trata_forms.infrastructure.sistema.caminhos import RAIZ
        with open(os.path.join(RAIZ, "cvc_bandeja.log"), "a",
                  encoding="utf-8") as f:
            f.write(_t.strftime("%d/%m %H:%M") + "  [layout] " + msg + "\n")
    except Exception:
        pass


def _forcar_layout_amplo(page, log=print):
    """Forca um viewport LARGO e fixo (1600x900), independente do tamanho FISICO
    da tela, p/ o OWA e o Jira renderizarem o layout DESKTOP ate num notebook de
    tela pequena. Sem isto, telas menores mudam o layout responsivo e os cliques
    / o cursor (Control+Home) caem em posicao diferente -> falha no botao Enviar
    e no preenchimento do encaminhamento. Via CDP. Nunca lanca.
    Pode ser desligado p/ teste com a variavel CVC_VIEWPORT_OVERRIDE=off."""
    if os.environ.get("CVC_VIEWPORT_OVERRIDE", "").strip().lower() == "off":
        _log_layout("override DESLIGADO (CVC_VIEWPORT_OVERRIDE=off)")
        return
    try:
        cdp = page.context.new_cdp_session(page)
        cdp.send("Emulation.setDeviceMetricsOverride", {
            "width": 1600, "height": 900,
            "deviceScaleFactor": 1, "mobile": False,
        })
        try:
            w = page.evaluate("window.innerWidth")
        except Exception:
            w = "?"
        _log_layout(f"viewport forcado p/ 1600x900 -> innerWidth={w}")
    except Exception as e:
        _log_layout(f"FALHOU ao forcar o viewport: {e}")


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
        # janela grande por padrao; CVC_WINDOW_SIZE=1024,600 forca uma pequena
        # (p/ TESTE de tela de notebook). Ex.: CVC_WINDOW_SIZE=1024,600
        (f"--window-size={os.environ['CVC_WINDOW_SIZE'].strip()}"
         if os.environ.get("CVC_WINDOW_SIZE", "").strip()
         else "--start-maximized"),
        url_inicial,
    ]
    # CREATE_BREAKAWAY_FROM_JOB: faz o Edge sobreviver quando o script termina
    # (sem DETACHED_PROCESS, para a janela aparecer em primeiro plano).
    creationflags = 0
    if os.name == "nt":
        CREATE_NEW_PROCESS_GROUP = 0x00000200
        CREATE_BREAKAWAY_FROM_JOB = 0x01000000
        CREATE_NO_WINDOW = 0x08000000       # nao pisca console
        creationflags = (CREATE_NEW_PROCESS_GROUP | CREATE_BREAKAWAY_FROM_JOB
                         | CREATE_NO_WINDOW)
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


def _avisar_senha_invalida(tentativa=1, total=3, log=print):
    """Aviso EXPLICITO (janela) de que a senha foi rejeitada e que as
    Configuracoes vao abrir p/ informar a correta. Na 1a vez sugere expiracao/
    alteracao; nas seguintes, alerta p/ CONFERIR A DIGITACAO. Sem isto o
    formulario abriria 'do nada' e o usuario nao saberia o porque."""
    if tentativa <= 1:
        motivo = ("A senha da conta do Outlook foi REJEITADA pela Microsoft "
                  "(provavelmente EXPIROU ou foi ALTERADA).")
    else:
        motivo = ("A senha informada AINDA está incorreta. "
                  "Confira a digitação (maiúsculas/minúsculas e espaços).")
    texto = (f"{motivo}\n\n"
             f"Tentativa {tentativa} de {total}.\n\n"
             "Vou abrir as Configurações para você informar a senha.\n"
             "Depois de salvar, o login continua automaticamente.")
    try:
        import tkinter as tk
        from tkinter import messagebox
        r = tk.Tk()
        r.withdraw()
        r.attributes("-topmost", True)
        messagebox.showwarning("Senha inválida ou expirada", texto, parent=r)
        r.destroy()
    except Exception as e:  # noqa: BLE001
        log(f"   (nao consegui exibir o aviso de senha: {e})")


def _encerrar_senha_invalida(popup_proc=None, log=print):
    """Mensagem AMIGAVEL final quando a senha nao foi validada apos varias
    tentativas: pede p/ atualizar a senha e tentar mais tarde, e encerra
    (fecha o fechador de popups). Sem stack trace / instrucoes tecnicas."""
    texto = ("Não foi possível validar a senha após várias tentativas.\n\n"
             "Verifique a senha correta da sua conta e abra o aplicativo "
             "novamente mais tarde para tentar de novo.")
    try:
        import tkinter as tk
        from tkinter import messagebox
        r = tk.Tk()
        r.withdraw()
        r.attributes("-topmost", True)
        messagebox.showerror("Login não concluído", texto, parent=r)
        r.destroy()
    except Exception as e:  # noqa: BLE001
        log(f"   (aviso final de senha: {e})")
    log("\n" + "!" * 70)
    log("  [ERRO] Senha nao validada apos varias tentativas. Atualize a senha "
        "e tente novamente mais tarde.")
    log("!" * 70)
    if popup_proc is not None:
        try:
            popup_proc.terminate()
        except Exception:
            pass


def _abrir_form_config_e_esperar(log=print):
    """Abre o formulario de configuracao (edita o config.xml) e ESPERA ele
    fechar. Frozen-aware: no .exe usa 'exe --config'; no codigo, 'python -m
    ...form_config'. Usado quando a senha salva foi rejeitada, p/ o usuario
    informar a senha nova sem sair do fluxo."""
    try:
        if getattr(sys, "frozen", False):
            cmd = [sys.executable, "--config"]
        else:
            cmd = [sys.executable, "-m",
                   "cvc_trata_forms.presentation.form_config"]
        subprocess.Popen(cmd, creationflags=0x08000000).wait()   # espera fechar
        return True
    except Exception as e:  # noqa: BLE001
        log(f"   (nao consegui abrir o formulario de config: {e})")
        return False


def _recarregar_credenciais(log=print):
    """Recarrega o config.xml do disco (o config_app cacheia os valores no
    import) e devolve (email, senha) ATUALIZADOS."""
    try:
        import importlib
        from cvc_trata_forms.infrastructure.config import config_app as _cfg
        importlib.reload(_cfg)
    except Exception as e:  # noqa: BLE001
        log(f"   (falha ao recarregar o config: {e})")
    return obter_outlook()


def _login_com_retry(login_fn, page, popup_proc, nome, log=print,
                     max_tentativas=3):
    """Executa login_fn(page, email, senha). Se a Microsoft REJEITAR a senha
    (CredencialInvalida), AVISA e abre o formulario de config para o usuario
    informar a senha correta, recarrega o config e tenta de novo. Repete ate
    'max_tentativas' (cobre erro de DIGITACAO). Esgotando as tentativas,
    ENCERRA com uma mensagem amigavel (atualize a senha e tente mais tarde).
    Retorna (ok, abortar). abortar=True quando desistiu."""
    email, senha = obter_outlook()
    for tentativa in range(1, max_tentativas + 1):
        try:
            ok = login_fn(page, email, senha, log=log)
            return ok, False
        except CredencialInvalida as e:
            log(f"  [SENHA] {e}  ({nome}) - tentativa {tentativa}/{max_tentativas}")
            if tentativa >= max_tentativas:
                _encerrar_senha_invalida(popup_proc, log)   # desiste (amigavel)
                return False, True
            _avisar_senha_invalida(tentativa, max_tentativas, log)
            _abrir_form_config_e_esperar(log)
            email, senha = _recarregar_credenciais(log)
            if not (email and senha):
                _encerrar_senha_invalida(popup_proc, log)
                return False, True
            log(f"   Tentando {nome} de novo p/ {email}...")
    return False, False


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
        if getattr(sys, "frozen", False):      # dentro do .exe (PyInstaller)
            cmd = [sys.executable, "--fechar-popups"]
        else:
            cmd = [sys.executable, "-m",
                   "cvc_trata_forms.infrastructure.sistema.fechar_popups", "--loop"]
        popup_proc = subprocess.Popen(cmd, creationflags=0x08000000)  # no window
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
    _forcar_layout_amplo(outlook)   # layout desktop mesmo em tela pequena
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
            ok_outlook, abortar = _login_com_retry(
                login_outlook, outlook, popup_proc, "OUTLOOK", log=print)
            if abortar:
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

    # Com a API do Jira ativa (usuario + token no config), NAO abrimos a aba
    # nem fazemos login SSO no Jira -- o chamado e criado/cancelado via REST.
    if JIRA_API_ATIVA:
        if popup_proc is not None:
            try:
                popup_proc.terminate()
            except Exception:
                pass
        print("\n" + "=" * 70)
        print("  PRONTO! Outlook logado. Jira sera via API (sem aba/login).")
        print("=" * 70)
        if INVISIVEL:
            esconder_edge()
        return

    # ---- PASSO 2: JIRA (so agora, com o Outlook ja logado) -----------------
    print("\n" + "-" * 70)
    print("  PASSO 2/2  -  JIRA  (Outlook ja logado, abrindo o Jira agora)")
    print("-" * 70)
    jira = garantir_aba(browser, ["atlassian"], URL_JIRA)
    _forcar_layout_amplo(jira)      # layout desktop mesmo em tela pequena

    # DUAS FRENTES tambem no Jira: automatico (SSO) com fallback manual.
    email, senha = obter_outlook()
    ok_jira = False
    if email and senha and not jira_logado(jira):
        print("   Tentando login automatico no Jira (SSO Microsoft)...")
        try:
            ok_jira, abortar = _login_com_retry(
                login_jira, jira, popup_proc, "JIRA", log=print)
            if abortar:
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
