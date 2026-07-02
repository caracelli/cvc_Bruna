"""
Login automatico no Outlook (Microsoft / Azure AD) com ESPERA DE MFA.

Fluxo:
  1. Preenche o e-mail e clica em "Avancar".
  2. Preenche a senha e clica em "Entrar".
  3. AGUARDA voce concluir o MFA (aprovar no celular / digitar codigo).
  4. Trata a tela "Continuar conectado?" (clica em Sim -> sessao dura mais).
  5. Confirma que chegou na caixa de e-mail.

Importante:
  - A senha vem do cofre do Windows (credenciais.py) -> nunca em texto puro.
  - O MFA continua sendo SEU (o script so espera).
  - Seletores sao os padroes do login.microsoftonline.com, com alternativas
    por tipo de campo (robustez). >>> Validar ao vivo na 1a execucao real.
"""

import time

# Seletores padrao da tela de login Microsoft (+ fallback por tipo)
SEL_EMAIL = "#i0116, input[type='email'][name='loginfmt']"
SEL_SENHA = "#i0118, input[type='password']"
SEL_BOTAO = "#idSIButton9, input[type='submit'], button[type='submit']"
SEL_KMSI_SIM = "#idSIButton9"          # "Sim" em "Continuar conectado?"
SEL_KMSI_NAO = "#idBtn_Back"           # "Nao" em "Continuar conectado?"

# sinais de que JA esta na caixa (logado)
POS_LOGADO = ["div[role='treeitem']", "div[role='option']",
              "div[role='navigation']"]
LOGIN_HOSTS = ("login.microsoftonline.com", "login.live.com",
               "login.microsoft.com")

# Texto do botao/opcao "Aprovar uma solicitacao" (MFA push) - MULTI-IDIOMA.
APROVAR_MFA = [
    "Aprovar uma solicitação", "Aprovar uma solicitacao",   # PT
    "Approve a request", "Send notification",               # EN
    "Aprobar una solicitud", "Aprueba una solicitud",       # ES
]


def loc_aprovar_mfa(page):
    """Retorna o locator do botao de aprovar MFA (qualquer idioma) ou None."""
    for t in APROVAR_MFA:
        try:
            loc = page.get_by_text(t, exact=False)
            if loc.count() > 0:
                return loc
        except Exception:
            pass
    return None


class CredencialInvalida(Exception):
    """E-mail ou senha rejeitados pela Microsoft."""


def _na_tela_login(page):
    u = page.url.lower()
    return any(h in u for h in LOGIN_HOSTS)


def _erro_credencial(page):
    """Detecta erro de e-mail/senha pela MS (IDs estaveis = multi-idioma).
    Retorna a mensagem do erro ou None."""
    try:
        if page.locator("#usernameError:visible").count() > 0:
            return ("E-MAIL/USUARIO invalido (a Microsoft nao reconheceu o "
                    "e-mail informado)")
        if page.locator("#passwordError:visible").count() > 0:
            return "SENHA invalida (a Microsoft rejeitou a senha)"
    except Exception:
        pass
    return None


def tratar_seletor_conta(page, email, log=print):
    """Trata a tela 'Escolha uma conta' (account picker).

    - Se houver um tile com o e-mail desejado -> clica nele.
    - Senao -> clica em 'Usar outra conta' (leva ao campo de e-mail).
    Retorna True se a tela existia e foi tratada; False se nao era essa tela.
    """
    txt_picker = ["Escolha uma conta", "Pick an account", "Choose an account",
                  "Escolha una cuenta", "Elegir una cuenta"]
    txt_outra = ["Use outra conta", "Usar outra conta", "Use another account",
                 "Usar otra cuenta", "Utilizar otra cuenta"]

    tem_picker = (
        page.locator("#otherTile, #otherTileText").count() > 0
        or any(page.get_by_text(t, exact=False).count() > 0 for t in txt_picker)
    )
    if not tem_picker:
        return False

    log("   Tela 'Escolha uma conta' detectada.")
    # 1) tile com o e-mail desejado
    tile = page.get_by_text(email, exact=False)
    if tile.count() > 0:
        log(f"   Selecionando a conta: {email}")
        tile.first.click()
        return True
    # 2) 'Usar outra conta'
    for t in txt_outra:
        op = page.get_by_text(t, exact=False)
        if op.count() > 0:
            log("   Conta desejada nao listada -> 'Usar outra conta'.")
            op.first.click()
            return True
    if page.locator("#otherTile").count() > 0:
        page.click("#otherTile")
        return True
    return False


def _ja_logado(page):
    if _na_tela_login(page):
        return False
    for sel in POS_LOGADO:
        try:
            if page.locator(sel).count() > 0:
                return True
        except Exception:
            pass
    return False


def preencher_credenciais(page, email, senha, log=print):
    """Etapas 1 e 2: digita e-mail e senha. Retorna True se enviou a senha."""
    # E-MAIL
    try:
        page.wait_for_selector(SEL_EMAIL, timeout=20000)
        page.fill(SEL_EMAIL, email)
        log("   E-mail preenchido. Avancando...")
        page.click(SEL_BOTAO)
    except Exception:
        log("   (campo de e-mail nao apareceu - talvez ja tenha passado)")

    # SENHA
    try:
        page.wait_for_selector(SEL_SENHA, timeout=20000)
        page.fill(SEL_SENHA, senha)
        log("   Senha preenchida. Entrando...")
        page.click(SEL_BOTAO)
        return True
    except Exception:
        log("   (campo de senha nao apareceu - verifique o fluxo)")
        return False


def aguardar_mfa_e_login(page, log=print, timeout_s=240, intervalo_s=2):
    """Etapas 3-5: espera o MFA, trata 'Continuar conectado?' e confirma."""
    log("   >> AGUARDANDO MFA: aprove no seu celular / digite o codigo.")
    fim = time.time() + timeout_s
    avisou_mfa = False
    while time.time() < fim:
        # ja chegou na caixa?
        if _ja_logado(page):
            log("   [OK] Login concluido - caixa de e-mail carregada.")
            return True

        # tela "Continuar conectado?" (KMSI) -> clicar Sim
        try:
            if page.locator("#KmsiCheckboxField, [name='DontShowAgain']").count() > 0 \
               or "kmsi" in page.url.lower():
                log("   Tela 'Continuar conectado?' - clicando em Sim.")
                page.click(SEL_KMSI_SIM)
                time.sleep(2)
                continue
        except Exception:
            pass

        if not avisou_mfa and _na_tela_login(page):
            avisou_mfa = True  # so loga a mensagem uma vez

        time.sleep(intervalo_s)

    log("   [X] Tempo esgotado aguardando MFA/login.")
    return False


def _existe(page, seletor):
    try:
        return page.locator(seletor).count() > 0
    except Exception:
        return False


def login_outlook(page, email, senha, log=print, timeout_mfa_s=240):
    """Login como MAQUINA DE ESTADOS: a cada ciclo observa a tela atual e
    age conforme o que estiver na frente (picker, e-mail, senha, KMSI, MFA).
    Bem mais robusto que passos com tempo fixo. Retorna True se logou.
    """
    if _ja_logado(page):
        log("   [OK] Outlook ja estava logado.")
        return True
    if not _na_tela_login(page):
        try:
            page.goto("https://outlook.office.com/mail/",
                      wait_until="domcontentloaded", timeout=30000)
        except Exception:
            pass   # a aba pode ja estar navegando (ERR_ABORTED) - segue o loop
        time.sleep(2)

    email_enviado = False
    senha_enviada = False
    avisou_mfa = False
    ultimo_aprovar = 0.0
    fim = time.time() + timeout_mfa_s + 60

    while time.time() < fim:
        # 1) chegou na caixa?
        if _ja_logado(page):
            log("   [OK] Login concluido - caixa de e-mail carregada.")
            return True

        # 1.5) credencial invalida? -> aborta com mensagem clara (nao fica preso)
        if email_enviado or senha_enviada:
            err = _erro_credencial(page)
            if err:
                raise CredencialInvalida(err)

        # 2) "Continuar conectado?" (KMSI) -> sempre NAO (nao persistir sessao)
        if _existe(page, "#KmsiCheckboxField") or "kmsi" in page.url.lower():
            log("   Tela 'Continuar conectado?' - clicando em NAO.")
            try:
                page.click(SEL_KMSI_NAO)
            except Exception:
                pass
            time.sleep(2)
            continue

        # 3) tela "Escolha uma conta"
        if tratar_seletor_conta(page, email, log=log):
            time.sleep(2.5)
            continue

        # 4) campo de e-mail
        if _existe(page, SEL_EMAIL) and not email_enviado:
            log("   Preenchendo e-mail...")
            try:
                page.fill(SEL_EMAIL, email)
                page.press(SEL_EMAIL, "Enter")   # Enter = envio confiavel
                email_enviado = True
            except Exception:
                pass
            time.sleep(2.5)
            continue

        # 5) campo de senha
        if _existe(page, SEL_SENHA) and not senha_enviada:
            log("   Preenchendo senha e enviando...")
            try:
                page.fill(SEL_SENHA, senha)
                page.press(SEL_SENHA, "Enter")   # Enter = envio confiavel
                senha_enviada = True
            except Exception:
                pass
            time.sleep(2.5)
            continue

        # 5.5) tela de MFA "Verifique sua identidade" / "Problema ao verificar"
        #      -> (re)disparar o push clicando em 'Aprovar uma solicitacao'
        if senha_enviada:
            aprovar = loc_aprovar_mfa(page)
            if aprovar is not None and (time.time() - ultimo_aprovar) > 25:
                from winutil import trazer_edge_frente
                trazer_edge_frente()
                log("   MFA: (re)enviando 'Aprovar uma solicitacao' (Authenticator)...")
                try:
                    aprovar.first.click()
                except Exception:
                    pass
                ultimo_aprovar = time.time()
                time.sleep(3)
                continue

        # 6) nada de campo -> provavelmente tela de MFA: aguardar voce
        if senha_enviada and not avisou_mfa:
            from winutil import trazer_edge_frente
            trazer_edge_frente()   # janela pra frente para ver o codigo
            log("   >> AGUARDANDO MFA: aprove no seu celular / digite o codigo.")
            avisou_mfa = True
        time.sleep(2)

    log("   [X] Tempo esgotado no login (picker/e-mail/senha/MFA).")
    return False
