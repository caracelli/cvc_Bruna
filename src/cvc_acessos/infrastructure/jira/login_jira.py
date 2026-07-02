"""
Login automatico no Jira Service Management (portal Service Desk) via SSO.

FLUXO REAL (observado nos prints), na ordem:
  1. "Service Desk - Insira o e-mail para entrar"  -> preenche e-mail -> PROXIMO
  2. "Service Desk - Use a conta da Atlassian"      -> "Continue com a conta
     da Atlassian"
  3. Modal do Jira "Entre para continuar"           -> clicar em "Microsoft"
     (NAO usar a senha da Atlassian: e diferente da Microsoft!)
  4. Microsoft "Escolha uma conta"                  -> seleciona -> logado.

KMSI ("Continuar conectado?") -> sempre NAO. A checkbox "Continuar conectado"
do modal do Jira fica desmarcada (nao mexemos nela).
"""

import time

from cvc_acessos.infrastructure.outlook.login_outlook import (
    tratar_seletor_conta, _existe, loc_aprovar_mfa, _erro_credencial,
    CredencialInvalida,
    SEL_EMAIL, SEL_SENHA, SEL_KMSI_NAO,
)

# hosts de login (se a URL contiver isto, NAO esta logado no Jira)
LOGIN_HOSTS_JIRA = ("login.microsoftonline.com", "login.live.com",
                    "login.microsoft.com", "id.atlassian.com", "/login")

# hosts de login da Microsoft (la sim usamos a senha MS)
MS_HOSTS = ("login.microsoftonline.com", "login.live.com",
            "login.microsoft.com")

# tela 1: Service Desk "Insira o e-mail" (campo + botao Proximo)
JSM_EMAIL = ("input[type='email'], #username, input[name='username'], "
             "input[name='email'], input[id*='email' i]")
# Botao "Proximo" do Service Desk (texto especifico p/ NAO colidir com o
# botao 'Continue com a conta da Atlassian', que tambem e submit).
JSM_PROXIMO = ("button:has-text('Próximo'), button:has-text('Proximo'), "
               "button:has-text('Next'), button:has-text('Siguiente'), "
               "input[type='submit'][value='Próximo'], "
               "input[type='submit'][value='Next'], "
               "input[type='submit'][value='Siguiente']")

# tela 2: "Continue com a conta da Atlassian" -- e um LINK (<a>), nao <button>.
# Aceita button/a/role=button com a frase ESPECIFICA de "conta Atlassian"
# (multi-idioma; nao pega o "Atlassian" solto do rodape).
BTN_ATLASSIAN_CONTINUE = (
    ":is(button,a,[role='button']):has-text('conta da Atlassian'), "    # PT
    ":is(button,a,[role='button']):has-text('Atlassian account'), "     # EN
    ":is(button,a,[role='button']):has-text('cuenta de Atlassian'), "   # ES
    ":is(button,a,[role='button']):has-text('cuenta Atlassian')"
)
# tela 3: opcao social Microsoft no modal do Jira
BTN_MICROSOFT_SSO = (
    "#microsoft, [data-testid*='microsoft' i], [id*='microsoft' i], "
    "button:has-text('Microsoft'), a:has-text('Microsoft')"
)


# URL da pagina de abertura de chamado (vem do config.xml)
from cvc_acessos.infrastructure.config.config_app import URL_JIRA as JIRA_CREATE_URL


def jira_logado(page):
    """PRONTO = NAO esta em tela de login E o FORMULARIO de chamado esta na
    tela (#summary). Evita falso positivo na tela da Microsoft."""
    u = page.url.lower()
    if any(h in u for h in LOGIN_HOSTS_JIRA):
        return False
    return (_existe(page, "#summary")
            or _existe(page, "input#request-type-select"))


def _atlassian_logado_sem_form(page):
    """Esta logado no Atlassian (host cvccorp), fora de tela de login, mas
    NAO no formulario (ex.: caiu na home). Sinal para navegar ao formulario."""
    u = page.url.lower()
    if "cvccorp.atlassian.net" not in u or "/login" in u:
        return False
    if _existe(page, "#summary") or _existe(page, "input#request-type-select"):
        return False
    # se ainda ha elementos de login na tela, NAO esta logado de fato
    if (_existe(page, JSM_EMAIL) or _existe(page, BTN_ATLASSIAN_CONTINUE)
            or _existe(page, BTN_MICROSOFT_SSO)):
        return False
    return True


def _on_ms(page):
    return any(h in page.url.lower() for h in MS_HOSTS)


def login_jira(page, email, senha, log=print, timeout_s=200,
               url_form=JIRA_CREATE_URL):
    """Loga no portal do Jira via SSO Microsoft. Retorna True ao chegar no
    FORMULARIO de abertura de chamado (navega ate ele se cair na home)."""
    if jira_logado(page):
        log("   [OK] Jira ja estava logado (no formulario).")
        return True

    jsm_email_enviado = False
    clicou_atlassian = False
    clicou_microsoft = False
    senha_ms_enviada = False
    ms_email_enviado = False
    avisou_mfa = False
    tentativas_form = 0
    ultimo_aprovar = 0.0
    fim = time.time() + timeout_s + 30

    while time.time() < fim:
        if jira_logado(page):
            log("   [OK] Jira logado - formulario do portal carregado.")
            return True

        # credencial invalida no SSO Microsoft -> aborta com mensagem clara
        if senha_ms_enviada or ms_email_enviado:
            err = _erro_credencial(page)
            if err:
                raise CredencialInvalida(err)

        # ---------- lado Microsoft ----------
        # KMSI -> sempre NAO
        if _existe(page, "#KmsiCheckboxField") or "kmsi" in page.url.lower():
            log("   Tela 'Continuar conectado?' - clicando em NAO.")
            try:
                page.click(SEL_KMSI_NAO)
            except Exception:
                pass
            time.sleep(2)
            continue

        if tratar_seletor_conta(page, email, log=log):
            time.sleep(2.5)
            continue

        # E-MAIL Microsoft ANTES da senha (a tela de e-mail vem primeiro)
        if _on_ms(page) and not ms_email_enviado and _existe(page, SEL_EMAIL):
            log("   (SSO) preenchendo e-mail Microsoft...")
            try:
                page.fill(SEL_EMAIL, email)
                page.press(SEL_EMAIL, "Enter")   # Enter = envio confiavel
                ms_email_enviado = True
            except Exception:
                pass
            time.sleep(2.5)
            continue

        # SENHA Microsoft
        if _on_ms(page) and not senha_ms_enviada and _existe(page, SEL_SENHA):
            log("   (SSO) preenchendo senha Microsoft e clicando Entrar...")
            try:
                page.fill(SEL_SENHA, senha)
                page.press(SEL_SENHA, "Enter")   # Enter = envio confiavel
                senha_ms_enviada = True
            except Exception:
                pass
            time.sleep(2.5)
            continue

        # ---------- lado Atlassian/Jira (so quando NAO esta no host MS) ----------
        # tela 1: Service Desk "Insira o e-mail" (e-mail + Proximo)
        if (not _on_ms(page) and not jsm_email_enviado
                and _existe(page, JSM_EMAIL) and _existe(page, JSM_PROXIMO)):
            log("   Jira: preenchendo e-mail do Service Desk e clicando Proximo...")
            try:
                page.fill(JSM_EMAIL, email)
                page.locator(JSM_PROXIMO).first.click()
                jsm_email_enviado = True
            except Exception:
                pass
            time.sleep(3)
            continue

        # tela 2: "Continue com a conta da Atlassian" (so no lado Atlassian!)
        if (not _on_ms(page) and not clicou_atlassian
                and _existe(page, BTN_ATLASSIAN_CONTINUE)):
            log("   Jira: 'Continue com a conta da Atlassian'...")
            try:
                page.locator(BTN_ATLASSIAN_CONTINUE).first.click()
                clicou_atlassian = True
            except Exception:
                pass
            time.sleep(3)
            continue

        # tela 3: SSO Microsoft (SO no lado Atlassian; nunca em tela MS, p/ nao
        # confundir com "Microsoft Authenticator" do MFA)
        if (not _on_ms(page) and not clicou_microsoft
                and _existe(page, BTN_MICROSOFT_SSO)):
            log("   Jira: entrando via SSO 'Microsoft'...")
            try:
                page.locator(BTN_MICROSOFT_SSO).first.click()
                clicou_microsoft = True
            except Exception:
                pass
            time.sleep(3)
            continue

        # logado no Atlassian mas caiu na HOME -> ir para a pagina do formulario
        if _atlassian_logado_sem_form(page) and tentativas_form < 3:
            log("   Jira logado; abrindo a pagina de abertura de chamado...")
            try:
                page.goto(url_form, wait_until="domcontentloaded", timeout=30000)
            except Exception:
                pass
            tentativas_form += 1
            time.sleep(3)
            continue

        # tela de MFA "Verifique sua identidade"/"Problema ao verificar" ->
        # (re)disparar o push clicando em 'Aprovar uma solicitacao'
        if senha_ms_enviada:
            aprovar = loc_aprovar_mfa(page)
            if aprovar is not None and (time.time() - ultimo_aprovar) > 25:
                from cvc_acessos.infrastructure.sistema.winutil import trazer_edge_frente
                trazer_edge_frente()
                log("   MFA: (re)enviando 'Aprovar uma solicitacao' (Authenticator)...")
                try:
                    aprovar.first.click()
                except Exception:
                    pass
                ultimo_aprovar = time.time()
                time.sleep(3)
                continue

        # nada reconhecido: pode ser a tela de MFA da Microsoft -> aguardar voce
        if senha_ms_enviada and not avisou_mfa:
            from cvc_acessos.infrastructure.sistema.winutil import trazer_edge_frente
            trazer_edge_frente()   # janela pra frente para ver o codigo
            log("   >> Se pedir MFA, aprove no celular / digite o codigo.")
            avisou_mfa = True
        time.sleep(2)

    log("   [X] Tempo esgotado no login do Jira.")
    return False
