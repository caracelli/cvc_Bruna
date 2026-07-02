"""
=============================================================================
 FLUXO PRINCIPAL  -  CVC / Gestao de Acessos
=============================================================================
Para cada e-mail do remetente "Microsoft Forms" NAO LIDO na caixa
compartilhada "Gestao de Acessos":

   1. extrai titulo / corpo / link "Exibir resultados"
   2. preenche o chamado no Jira (Resumen + Descripcion formatada + Tipo=Forms)
   3. [real] envia o chamado e captura o codigo (ex: GAAR-1234) + link
   4. [real] ENCAMINHA o e-mail pro grupo (assunto = "<codigo> - <titulo>",
             corpo com o nº do chamado clicavel) -> sem categoria, nao acumula
   5. [real] marca como lido
   6. [real] move o e-mail para a subpasta "Finalizados"

-----------------------------------------------------------------------------
 DRY-RUN (ensaio)
-----------------------------------------------------------------------------
 DRY_RUN = True  -> LE, EXTRAI e PREENCHE o formulario do Jira (preview),
                    mas NAO envia o chamado, NAO encaminha, NAO move e NAO
                    marca como lido. So mostra o que faria.
 DRY_RUN = False -> executa de verdade e, no fim, faz logoff.

 USO:
     python fluxo.py
=============================================================================
"""

import re
import sys
import time

from checar_login import checar
from credenciais import obter_outlook
from login_jira import login_jira, jira_logado
from outlook_web import abrir_inbox_compartilhada, ler_emails

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# =========================  CONFIGURACAO  ===================================
# Tudo configuravel vem do config.xml (via config_app).
from config_app import (
    DRY_RUN, CAIXA, REMETENTE_FILTRO, SO_NAO_LIDOS, SUBPASTA_DESTINO,
    MARCAR_COMO_LIDO, ENCERRAR_SESSOES_NO_FIM, URL_JIRA,
    JIRA_TIPO_SOLICITUD, JIRA_COR_TITULO, JIRA_CAT_PREFIXO,
    ENCAMINHAR_ATIVO, ENCAMINHAR_DESTINATARIOS, ENCAMINHAR_ASSUNTO,
    ENCAMINHAR_PREFIXO, ENCAMINHAR_SEPARADOR,
)

# prefixos do assunto Forms a remover p/ obter o titulo limpo (multi-idioma)
PREFIXOS = ["Nova resposta de ", "Nova resposta a ",
            "Nueva respuesta de ", "Nueva respuesta a ",
            "New response for ", "New response to "]
# ===========================================================================


# ---------------------------------------------------------------------------
#  EXTRACAO DO E-MAIL (somente leitura)
# ---------------------------------------------------------------------------
def extrair_email(outlook, indice):
    """Abre o e-mail no indice e devolve {titulo, corpo, link}."""
    outlook.locator("div[role='option']").nth(indice).click()
    time.sleep(3.5)

    # titulo limpo (do cabecalho do assunto)
    raw = ""
    hs = outlook.locator("[role='heading']")
    for i in range(hs.count()):
        t = (hs.nth(i).inner_text() or "").strip()
        low = t.lower()
        if "resposta" in low or "respuesta" in low or "response" in low:
            raw = t
            break
    for corte in ["Resumir", "Summarize", "\n"]:
        raw = raw.split(corte)[0].strip()
    for pre in PREFIXOS:
        if raw.lower().startswith(pre.lower()):
            raw = raw[len(pre):].strip()
            break
    titulo = raw

    # corpo (linha principal) - elemento mais especifico (menor texto)
    corpo = ""
    for frag in ["recebeu uma nova resposta", "recibió una nueva respuesta",
                 "received a new response"]:
        loc = outlook.get_by_text(frag, exact=False)
        n = loc.count()
        if n > 0:
            melhor = None
            for i in range(n):
                try:
                    t = loc.nth(i).inner_text().strip().replace("\n", " ")
                except Exception:
                    continue
                if t and (melhor is None or len(t) < len(melhor)):
                    melhor = t
            corpo = melhor or ""
            break

    # link do "Exibir resultados"
    link = ""
    a = outlook.locator("a:has-text('Exibir resultados'), "
                        "a:has-text('Ver resultados'), a:has-text('View results')")
    if a.count() > 0:
        link = a.first.get_attribute("href") or ""

    return {"titulo": titulo, "corpo": corpo, "link": link}


# ---------------------------------------------------------------------------
#  PREENCHIMENTO DO FORMULARIO DO JIRA (preview - reversivel, sem enviar)
# ---------------------------------------------------------------------------
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


# ---------------------------------------------------------------------------
#  ACOES REAIS (so com DRY_RUN = False)  ->  VALIDAR AO VIVO antes do go-live
# ---------------------------------------------------------------------------
def enviar_e_capturar_codigo(jira):
    """Clica em Enviar e captura o codigo do chamado (VALIDADO ao vivo: GAAR-4).
    Captura pela URL do chamado criado (.../portal/<n>/<CHAVE>?created=true),
    que e o jeito mais confiavel; cai no corpo da pagina como fallback."""
    url_antes = jira.url
    jira.locator(
        "button:has-text('Enviar'), button:has-text('Crear'), "
        "button:has-text('Criar'), button:has-text('Send'), button[type='submit']"
    ).first.click()
    time.sleep(5)
    # 1) chave pela URL do chamado criado (ex.: /portal/1984/GAAR-4)
    m = re.search(r"/portal/\d+/([A-Z]{2,8}-\d+)", jira.url)
    if m and jira.url != url_antes:
        return m.group(1)
    # 2) fallback: procura no corpo da pagina
    txt = jira.locator("body").inner_text()
    m2 = re.search(r"\b[A-Z]{2,8}-\d+\b", txt)
    return m2.group(0) if m2 else "??-?"


def _abrir_menu_contexto(outlook, indice, item_nome, retries=4):
    """Abre o menu de contexto (SO botao direito - nao abre/le o e-mail) e
    posiciona em 'item_nome'. Com RETRY: o menu as vezes nao abre logo apos
    uma movimentacao."""
    for _ in range(retries):
        try:
            outlook.keyboard.press("Escape")
            time.sleep(0.4)
            row = outlook.locator("div[role='option']").nth(indice)
            row.click(button="right")
            time.sleep(1.2)
            mi = outlook.get_by_role("menuitem", name=item_nome, exact=False)
            if mi.count() > 0 and mi.first.is_visible():
                mi.first.hover()
                time.sleep(1.0)
                return True
        except Exception:
            time.sleep(0.6)
    return False


def aplicar_categoria(outlook, indice, codigo):
    """Aplica a categoria 'codigo' no e-mail (VALIDADO ao vivo).
    Categorizar -> se ja existe (menuitemcheckbox) marca; senao Nova categoria."""
    if not _abrir_menu_contexto(outlook, indice, "Categorizar"):
        raise RuntimeError("menu 'Categorizar' nao abriu")
    cb = outlook.get_by_role("menuitemcheckbox", name=codigo, exact=False)
    if cb.count() > 0:
        if cb.first.get_attribute("aria-checked") != "true":
            cb.first.click()
    else:
        outlook.get_by_role("menuitem", name="Nova categoria",
                            exact=False).first.click()
        time.sleep(1.5)
        outlook.fill("input[placeholder='Nomeie sua categoria']", codigo)
        time.sleep(0.5)
        outlook.get_by_role("button", name="Salvar", exact=False).first.click()
    time.sleep(1)


def mover_email(outlook, indice, pasta):
    """Move o e-mail para a subpasta 'pasta' (VALIDADO ao vivo).
    Mover -> Selecionar uma pasta diferente -> busca -> clica o .fui-TreeItemLayout."""
    if not _abrir_menu_contexto(outlook, indice, "Mover"):
        raise RuntimeError("menu 'Mover' nao abriu")
    outlook.get_by_role("menuitem", name="Selecionar uma pasta diferente",
                        exact=False).first.click()
    time.sleep(2)
    outlook.fill("input[placeholder='Digite o nome da pasta ou do grupo']", pasta)
    time.sleep(2)
    tt = outlook.locator("[role='dialog'] [role='treeitem']")
    alvo = None
    for i in range(tt.count()):
        t = (tt.nth(i).inner_text() or "").lower()
        # a pasta destino (nao a raiz 'Gestao de Acessos')
        if pasta.lower() in t and "gest" not in t:
            alvo = tt.nth(i)
            break
    if alvo is None:
        raise RuntimeError(f"Pasta '{pasta}' nao encontrada no dialogo de mover.")
    alvo.locator(".fui-TreeItemLayout").first.click()
    time.sleep(0.8)
    outlook.get_by_role("button", name="Mover", exact=True).first.click()
    time.sleep(3)


def marcar_lido(outlook, indice):
    """Marca o e-mail como lido (botao de alternancia da linha)."""
    row = outlook.locator("div[role='option']").nth(indice)
    btn = row.locator("[title='Marcar como lido'], [title='Mark as read']")
    if btn.count() > 0:
        btn.first.click()


def _inserir_url_no_dialogo(outlook, link):
    """No dialogo 'Inserir link' (Ctrl+K), preenche o campo de URL (#linkInput)
    e clica OK. O 'Exibir como' ja vem com o texto selecionado (o ticket)."""
    time.sleep(1.2)
    campo = outlook.locator("#linkInput")
    if campo.count() == 0:
        campo = outlook.locator(
            "input[aria-label*='ndere' i], input[placeholder*='URL' i]")
    if campo.count() > 0:
        campo.first.fill(link)   # fill nao precisa de clique (evita backdrop)
        time.sleep(0.3)
    ok = outlook.get_by_role("button", name="OK", exact=True)
    if ok.count() > 0 and ok.first.is_visible():
        ok.first.click()
    time.sleep(0.6)


def _confirmar_popup_enviar(outlook):
    """Trata o popup 'parece que esqueceu de anexar' (botoes 'Enviar' /
    'Nao enviar') que aparece ao enviar. Clica 'Enviar' se surgir."""
    time.sleep(1.5)
    dlg = outlook.locator("[role='dialog'], [role='alertdialog']")
    if dlg.count() == 0:
        return
    b = outlook.get_by_role("button", name="Enviar", exact=True)
    for i in range(b.count()):
        try:
            if b.nth(i).is_visible():
                b.nth(i).click(); return
        except Exception:
            pass


def encaminhar_email(outlook, indice, destinatarios, assunto, ticket, link,
                     prefixo="Chamado aberto: ",
                     separador="--- Mensagem original (Microsoft Forms) abaixo ---"):
    """Encaminha o e-mail (indice) para 'destinatarios' com o assunto dado e,
    no TOPO do corpo, a linha: '<prefixo><ticket como HYPERLINK p/ link>' +
    o separador. Retorna True se enviou. VALIDADO ao vivo.

    Compose OWA: Para=[contenteditable][aria-label='Para'];
    Assunto=input[aria-label='Assunto']; Corpo=[role='textbox'][aria-label=
    'Corpo da mensagem']; link via Ctrl+K (Exibir como=texto sel., URL=campo).
    O 'De' sai automaticamente como a caixa compartilhada. Ao enviar, pode
    surgir o popup de anexos -> _confirmar_popup_enviar clica 'Enviar'."""
    if not _abrir_menu_contexto(outlook, indice, "Encaminhar"):
        raise RuntimeError("menu 'Encaminhar' nao abriu")
    outlook.get_by_role("menuitem", name="Encaminhar", exact=False).first.click()
    time.sleep(3.5)

    # Para (people picker)
    para = outlook.locator("[contenteditable='true'][aria-label='Para']").first
    para.click()
    time.sleep(0.4)
    for dest in destinatarios:
        outlook.keyboard.type(dest)
        time.sleep(0.9)
        outlook.keyboard.press("Enter")   # resolve o chip do destinatario
        time.sleep(0.6)

    # Assunto (substitui o "Enc: ...")
    subj = outlook.locator("input[aria-label='Assunto']").first
    subj.click()
    time.sleep(0.3)
    outlook.keyboard.press("Control+a")
    outlook.keyboard.press("Delete")
    time.sleep(0.2)
    outlook.keyboard.type(assunto)

    # Corpo: prefixo + numero do chamado como HYPERLINK, no TOPO
    corpo = outlook.locator(
        "[role='textbox'][aria-label='Corpo da mensagem']").first
    corpo.click()
    time.sleep(0.3)
    outlook.keyboard.press("Control+Home")
    time.sleep(0.2)
    # 1) escreve o cabecalho como TEXTO PLANO. O OWA COLAPSA paragrafos vazios
    #    na exibicao -> a linha em branco leva um espaco nao-quebravel ( )
    #    pra nao sumir.
    def _linha_branco():
        outlook.keyboard.press("Enter"); time.sleep(0.15)
        outlook.keyboard.insert_text("\u00a0"); time.sleep(0.1)  # nbsp (nao colapsa)
        outlook.keyboard.press("Enter"); time.sleep(0.15)

    # garante 1 espaco entre o prefixo e o numero (o config faz strip)
    pref = (prefixo.rstrip() + " ") if prefixo.strip() else ""
    outlook.keyboard.type(pref)
    outlook.keyboard.type(ticket)
    _linha_branco()                            # linha em branco (nbsp)
    if separador:
        outlook.keyboard.type(separador)
        _linha_branco()                        # linha em branco (nbsp)
    time.sleep(0.3)
    # 2) volta e transforma SO o ticket em HYPERLINK (por ultimo, pra nao
    #    depender de digitar nada depois de fechar o dialogo do link)
    outlook.keyboard.press("Control+Home")
    for _ in range(len(pref)):
        outlook.keyboard.press("ArrowRight")
    for _ in range(len(ticket)):
        outlook.keyboard.press("Shift+ArrowRight")
    time.sleep(0.3)
    outlook.keyboard.press("Control+k")       # abre 'Inserir link'
    _inserir_url_no_dialogo(outlook, link)
    time.sleep(0.4)

    # Enviar (+ trata popup de anexos)
    outlook.get_by_role("button", name="Enviar", exact=False).first.click()
    _confirmar_popup_enviar(outlook)
    time.sleep(3)
    return True


# ---------------------------------------------------------------------------
#  SESSAO
# ---------------------------------------------------------------------------
def _garantir_sessoes():
    est = checar()
    browser = est["browser"]
    outlook = est["page_outlook"]
    jira = est["page_jira"]
    if not est["outlook"].get("ok"):
        raise RuntimeError(
            "Outlook nao esta logado. Rode 'python iniciar_sessoes.py' antes."
        )
    if not est["jira"].get("ok"):
        print("   Jira pediu re-login (SSO) - resolvendo...")
        email, senha = obter_outlook()
        if jira is None:
            ctx = browser.contexts[0]
            jira = ctx.new_page()
            jira.goto(URL_JIRA, wait_until="domcontentloaded")
        if not login_jira(jira, email, senha, log=print):
            raise RuntimeError("Nao consegui religar o Jira.")
    return browser, outlook, jira


# ---------------------------------------------------------------------------
#  ORQUESTRACAO
# ---------------------------------------------------------------------------
def _tem_cat_jira(email):
    """True se o e-mail ja tem uma categoria 'JIRA -*' (= ja processado)."""
    pref = JIRA_CAT_PREFIXO.lower()
    return any(c.lower().startswith(pref) for c in email.get("categorias", []))


def _filtrar(outlook):
    termo = REMETENTE_FILTRO.lower()
    return [
        e for e in ler_emails(outlook, 100)
        if (not termo or termo in e["resumo"].lower())
        and (not SO_NAO_LIDOS or e["nao_lido"])
        and not _tem_cat_jira(e)        # pula os que ja tem categoria JIRA -*
    ]


def _plano(n, dados):
    print(f"\n  --- E-MAIL {n} ---")
    print(f"    Resumen     : {dados['titulo']}")
    print(f"    Corpo       : {dados['corpo']}")
    print(f"    Link        : {dados['link'][:70]}...")
    print(f"    Tipo        : Forms")


def processar(outlook, jira):
    """UM ciclo: abre a caixa, filtra e processa. NAO faz login nem logoff
    (quem chama cuida disso). Retorna (encontrados, processados)."""
    aberto = abrir_inbox_compartilhada(outlook, CAIXA)
    print(f">> Caixa de Entrada: {aberto}")

    if DRY_RUN:
        # ENSAIO: itera a lista (nao move), preenche o form e mostra o plano
        alvo = _filtrar(outlook)
        print(f">> {len(alvo)} e-mail(s) Microsoft Forms NAO lidos.\n")
        for n, e in enumerate(alvo, 1):
            dados = extrair_email(outlook, e["indice"])
            _plano(n, dados)
            jira.bring_to_front()
            ok = preencher_formulario(jira, dados)
            print(f"    Form Jira   : {'preenchido (preview)' if ok else 'FALHOU'}")
            print("    [DRY-RUN] enviaria o chamado + pegaria o codigo")
            if ENCAMINHAR_ATIVO and ENCAMINHAR_DESTINATARIOS:
                print(f"    [DRY-RUN] encaminharia p/ {ENCAMINHAR_DESTINATARIOS} "
                      f"(assunto '<ticket> - {dados['titulo']}', nº clicavel)")
            print(f"    [DRY-RUN] mover p/ '{SUBPASTA_DESTINO}' + marcar lido")
        print("\n>> DRY-RUN: nada enviado/movido/alterado.")
        return len(alvo), 0

    # EXECUCAO REAL: processa o 1o da fila ate acabar (mover tira da lista)
    encontrados = len(_filtrar(outlook))
    processados = 0
    while True:
        alvo = _filtrar(outlook)
        if not alvo:
            break
        e = alvo[0]
        dados = extrair_email(outlook, e["indice"])
        print(f"\n>> Processando: {dados['titulo']}")
        jira.bring_to_front()
        if not preencher_formulario(jira, dados):
            print("   [ERRO] nao preencheu o formulario; parando.")
            break
        codigo = enviar_e_capturar_codigo(jira)
        link = jira.url.split("?")[0]                # URL do chamado criado
        print(f"   Chamado criado: {codigo}  ({link})")
        outlook.bring_to_front()
        # NOTIFICA o grupo: encaminha com o nº do chamado no assunto + link
        # clicavel no corpo (sem categoria -> nao acumula lista-mestre).
        if ENCAMINHAR_ATIVO and ENCAMINHAR_DESTINATARIOS:
            assunto = ENCAMINHAR_ASSUNTO.format(ticket=codigo,
                                                assunto=dados["titulo"])
            encaminhar_email(outlook, e["indice"], ENCAMINHAR_DESTINATARIOS,
                             assunto, codigo, link, ENCAMINHAR_PREFIXO,
                             ENCAMINHAR_SEPARADOR)
            print(f"   Encaminhado p/ {len(ENCAMINHAR_DESTINATARIOS)} "
                  f"destinatario(s).")
        if MARCAR_COMO_LIDO:
            marcar_lido(outlook, e["indice"])
        mover_email(outlook, e["indice"], SUBPASTA_DESTINO)
        processados += 1
        jira.goto(URL_JIRA, wait_until="domcontentloaded")
        time.sleep(2)

    print(f"\n>> {processados} e-mail(s) processado(s).")
    return encontrados, processados


def main():
    modo = "[ DRY-RUN / ENSAIO ]" if DRY_RUN else "[ EXECUCAO REAL ]"
    print("=" * 74)
    print(f"  FLUXO CVC / GESTAO DE ACESSOS   {modo}")
    print("=" * 74)

    print(">> Verificando sessoes...")
    browser, outlook, jira = _garantir_sessoes()
    print("   OK: Outlook e Jira logados.\n")

    processar(outlook, jira)

    if not DRY_RUN and ENCERRAR_SESSOES_NO_FIM:
        from encerrar_sessoes import logoff_completo
        print(">> Logoff (Outlook + Jira)...")
        logoff_completo(browser, log=print)


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:  # noqa: BLE001
        print("\n[ERRO] " + str(ex))
        raise SystemExit(1)
