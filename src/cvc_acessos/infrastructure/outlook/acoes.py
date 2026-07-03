"""
Infra / Outlook — ACOES no e-mail (adapters que manipulam o DOM do OWA):
extrair, categorizar, mover, marcar lido e ENCAMINHAR (com hyperlink).
Recebem sempre a 'page' do Outlook (nao criam sessao).
"""

import time

from cvc_acessos.domain.regras import limpar_titulo


def extrair_email(outlook, indice):
    """Abre o e-mail no indice e devolve {titulo, corpo, link}."""
    outlook.locator("div[role='option']").nth(indice).click()
    time.sleep(3.5)

    # titulo (do cabecalho do assunto) -> limpo pela camada de dominio
    raw = ""
    hs = outlook.locator("[role='heading']")
    for i in range(hs.count()):
        t = (hs.nth(i).inner_text() or "").strip()
        low = t.lower()
        if "resposta" in low or "respuesta" in low or "response" in low:
            raw = t
            break
    titulo = limpar_titulo(raw)

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


def marcar_nao_lido(outlook, indice):
    """Marca o e-mail como NAO lido (usado p/ restaurar no DRY-RUN)."""
    row = outlook.locator("div[role='option']").nth(indice)
    btn = row.locator(
        "[title='Marcar como não lido'], [title='Mark as unread']")
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
                b.nth(i).click()
                return
        except Exception:
            pass


def encaminhar_email(outlook, indice, destinatarios, assunto, ticket, link,
                     prefixo="Chamado aberto: ",
                     separador="--- Mensagem original (Microsoft Forms) abaixo ---"):
    """Encaminha o e-mail (indice) para 'destinatarios' com o assunto dado e,
    no TOPO do corpo, a linha '<prefixo><ticket como HYPERLINK p/ link>' + o
    separador. Retorna True se enviou. VALIDADO ao vivo.

    Compose OWA: Para=[contenteditable][aria-label='Para'];
    Assunto=input[aria-label='Assunto']; Corpo=[role='textbox'][aria-label=
    'Corpo da mensagem']; link via Ctrl+K (Exibir como=texto sel., URL=#linkInput).
    O 'De' sai automaticamente como a caixa compartilhada. Ao enviar pode
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

    # O OWA COLAPSA paragrafos vazios -> a linha em branco leva um nbsp.
    def _linha_branco():
        outlook.keyboard.press("Enter")
        time.sleep(0.15)
        outlook.keyboard.insert_text(" ")
        time.sleep(0.1)
        outlook.keyboard.press("Enter")
        time.sleep(0.15)

    # garante 1 espaco entre o prefixo e o numero (o config faz strip)
    pref = (prefixo.rstrip() + " ") if prefixo.strip() else ""
    outlook.keyboard.type(pref)
    outlook.keyboard.type(ticket)
    _linha_branco()
    if separador:
        outlook.keyboard.type(separador)
        _linha_branco()
    time.sleep(0.3)

    # transforma SO o ticket em HYPERLINK (por ultimo, pra nao depender de
    # digitar nada apos fechar o dialogo do link)
    outlook.keyboard.press("Control+Home")
    for _ in range(len(pref)):
        outlook.keyboard.press("ArrowRight")
    for _ in range(len(ticket)):
        outlook.keyboard.press("Shift+ArrowRight")
    time.sleep(0.3)
    outlook.keyboard.press("Control+k")
    _inserir_url_no_dialogo(outlook, link)
    time.sleep(0.4)

    # Enviar (+ trata popup de anexos)
    outlook.get_by_role("button", name="Enviar", exact=False).first.click()
    _confirmar_popup_enviar(outlook)
    time.sleep(3)
    return True
