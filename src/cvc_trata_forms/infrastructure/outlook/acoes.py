"""
Infra / Outlook — ACOES no e-mail (adapters que manipulam o DOM do OWA):
extrair, categorizar, mover, marcar lido e ENCAMINHAR (com hyperlink).
Recebem sempre a 'page' do Outlook (nao criam sessao).
"""

import time

from cvc_trata_forms.domain.regras import limpar_titulo
from cvc_trata_forms.infrastructure.sistema.diagnostico import salvar_diagnostico


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


# pastas a EXCLUIR ao voltar pro Inbox (evita pasta pessoal/lixo/root)
_EXCL_INBOX = ("finalizados", "gest", "junk", "caixa", "draft", "sent",
               "deleted", "exclu", "archive", "arquivo", "notes",
               "conversation", "rascunho", "enviado", "lixo", "morto",
               "histórico", "historico")


def mover_para_inbox(outlook, indice, log=print):
    """Move o e-mail de volta pro Inbox da caixa COMPARTILHADA (usado no DEMO
    p/ desfazer). Busca 'Inbox' no dialogo e clica o treeitem que casa 'inbox'
    e NAO casa nenhuma pasta pessoal/lixo/root (_EXCL_INBOX).
    Retorna True se moveu; False se nao conseguiu (NUNCA lanca -> nao derruba o
    ciclo). Em caso de falha, salva um diagnostico (print+HTML) do estado atual
    p/ inspecao remota (o menu 'Mover' se comporta diferente na subpasta)."""
    try:
        if not _abrir_menu_contexto(outlook, indice, "Mover"):
            salvar_diagnostico(outlook, "mover_inbox_menu_nao_abriu", log)
            return False
        # timeout curto (8s): se nao aparecer, capturamos o menu em vez de
        # pendurar 30s no default do Playwright.
        outlook.get_by_role("menuitem", name="Selecionar uma pasta diferente",
                            exact=False).first.click(timeout=8000)
        time.sleep(2)
        outlook.fill("input[placeholder='Digite o nome da pasta ou do grupo']",
                     "Inbox")
        time.sleep(2)
        tt = outlook.locator("[role='dialog'] [role='treeitem']")
        for i in range(tt.count()):
            t = (tt.nth(i).inner_text() or "").lower()
            if "inbox" in t and all(x not in t for x in _EXCL_INBOX):
                tt.nth(i).locator(".fui-TreeItemLayout").first.click()
                time.sleep(0.8)
                outlook.get_by_role("button", name="Mover", exact=True).first.click()
                time.sleep(3)
                return True
        # menu abriu mas nao achamos o 'Inbox' na arvore -> diagnostica
        log("   [mover_inbox] 'Inbox' nao encontrado na arvore do dialogo.")
        salvar_diagnostico(outlook, "mover_inbox_sem_inbox_na_arvore", log)
        try:
            outlook.keyboard.press("Escape")
        except Exception:
            pass
        return False
    except Exception as e:
        log(f"   [mover_inbox] FALHA ao mover de volta: {e}")
        salvar_diagnostico(outlook, "mover_inbox_falhou", log)
        try:
            outlook.keyboard.press("Escape")
        except Exception:
            pass
        return False


def _toggle_leitura(outlook, indice, titulos):
    """Clica o botao de alternancia lido/nao-lido da LINHA. O botao as vezes
    so aparece no HOVER -> passa o mouse antes."""
    row = outlook.locator("div[role='option']").nth(indice)
    try:
        row.scroll_into_view_if_needed(timeout=2000)
    except Exception:
        pass
    try:
        row.hover()
        time.sleep(0.3)
    except Exception:
        pass
    btn = row.locator(titulos)
    if btn.count() > 0:
        try:
            btn.first.click()
            return True
        except Exception:
            pass
    return False


def marcar_lido(outlook, indice):
    """Marca o e-mail como lido (botao de alternancia da linha)."""
    return _toggle_leitura(
        outlook, indice,
        "[title='Marcar como lido'], [title='Mark as read'], "
        "button[aria-label*='como lido' i], button[aria-label*='as read' i]")


def marcar_nao_lido(outlook, indice):
    """Marca o e-mail como NAO lido (restaura o estado no DRY-RUN/teste)."""
    return _toggle_leitura(
        outlook, indice,
        "[title='Marcar como não lido'], [title='Mark as unread'], "
        "button[aria-label*='não lido' i], button[aria-label*='as unread' i]")


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


def _selecionar_ticket_no_corpo(corpo, ticket):
    """Seleciona EXATAMENTE o texto do ticket dentro do corpo, via Range do DOM.

    Antes a selecao era por teclado (Control+Home + N setas p/ pular o prefixo
    + N com Shift). Quando o Control+Home NAO levava o cursor pro inicio do
    corpo (acontece no OWA conforme o foco/layout), as setas andavam a partir
    do lugar errado e o link caia numa palavra qualquer do separador (ex.:
    'riginal', de 'Mensagem original'). Aqui a posicao vem do TEXTO, nao de
    contagem de teclas. Retorna True se achou e selecionou o ticket."""
    return bool(corpo.evaluate(
        """(el, alvo) => {
            // mapeia os nos de texto do corpo com o offset acumulado, pra
            // achar o ticket mesmo se ele estiver quebrado em varios nos
            const nos = [];
            const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
            let n, texto = "";
            while ((n = w.nextNode())) {
                nos.push([n, texto.length]);
                texto += n.nodeValue;
            }
            const ini = texto.indexOf(alvo);   // 1a ocorrencia = nossa linha
            if (ini < 0) return false;
            const fim = ini + alvo.length;
            let noI = null, offI = 0, noF = null, offF = 0;
            for (const [no, base] of nos) {
                const len = no.nodeValue.length;
                if (noI === null && ini >= base && ini < base + len) {
                    noI = no; offI = ini - base;
                }
                if (noF === null && fim > base && fim <= base + len) {
                    noF = no; offF = fim - base;
                }
            }
            if (!noI || !noF) return false;
            const r = document.createRange();
            r.setStart(noI, offI);
            r.setEnd(noF, offF);
            const s = window.getSelection();
            s.removeAllRanges();
            s.addRange(r);
            return true;
        }""", ticket))


def _conferir_link_do_ticket(corpo, ticket, link):
    """Confere no DOM se o hyperlink ficou no numero do chamado.
    {ok: existe <a> cujo texto e o ticket, errado: [textos dos <a> que apontam
    pro link do chamado mas NAO sao o ticket]} - 'errado' e o sintoma da falha
    antiga (link na palavra do separador)."""
    return corpo.evaluate(
        """(el, args) => {
            const as = Array.from(el.querySelectorAll('a'));
            const txt = (a) => (a.textContent || '').trim();
            return {
                ok: as.some(a => txt(a) === args.ticket),
                errado: as.filter(a => txt(a) !== args.ticket &&
                                       (a.getAttribute('href') || '') === args.link)
                          .map(txt).slice(0, 3),
            };
        }""", {"ticket": ticket, "link": link})


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
                     separador="--- Mensagem original (Microsoft Forms) abaixo ---",
                     log=print):
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

    # Se o campo 'Para' do compose nao aparecer, algo saiu do esperado (compose
    # nao abriu / layout diferente) -> captura print+HTML e aborta com clareza,
    # em vez de digitar no lugar errado e mandar e-mail torto.
    if outlook.locator("[contenteditable='true'][aria-label='Para']").count() == 0:
        log("   [encaminhar] campo 'Para' NAO apareceu (compose nao abriu?).")
        salvar_diagnostico(outlook, "encaminhar_compose_nao_abriu", log)
        raise RuntimeError("compose de Encaminhar nao abriu.")

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
    # digitar nada apos fechar o dialogo do link). A selecao vem do TEXTO
    # (Range no DOM) - contagem de setas errava a posicao no cliente.
    corpo.click()          # foco no corpo (o Ctrl+K age na selecao do editor)
    time.sleep(0.2)
    if not _selecionar_ticket_no_corpo(corpo, ticket):
        log(f"   [encaminhar] nao achei '{ticket}' no corpo p/ aplicar o link; "
            "segue com o numero em texto puro.")
        salvar_diagnostico(outlook, "encaminhar_ticket_nao_achado", log)
    else:
        outlook.keyboard.press("Control+k")
        _inserir_url_no_dialogo(outlook, link)
        time.sleep(0.4)
        conf = _conferir_link_do_ticket(corpo, ticket, link)
        if conf.get("errado"):
            # link colou em outra palavra (bug 'riginal') -> NAO manda torto
            log(f"   [encaminhar] o link caiu em {conf['errado']} em vez de "
                f"'{ticket}'. Abortando o envio.")
            salvar_diagnostico(outlook, "encaminhar_link_palavra_errada", log)
            raise RuntimeError(
                f"hyperlink aplicado no texto errado ({conf['errado']}), "
                f"esperado '{ticket}'.")
        if not conf.get("ok"):
            log(f"   [encaminhar] o numero '{ticket}' ficou sem hyperlink "
                "(dialogo do link nao aplicou); envio segue mesmo assim.")
            salvar_diagnostico(outlook, "encaminhar_link_nao_aplicado", log)

    # Enviar (+ trata popup de anexos)
    outlook.get_by_role("button", name="Enviar", exact=False).first.click()
    _confirmar_popup_enviar(outlook)
    time.sleep(3)
    return True
