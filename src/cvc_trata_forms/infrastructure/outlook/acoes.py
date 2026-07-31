"""
Infra / Outlook — ACOES no e-mail (adapters que manipulam o DOM do OWA):
extrair, categorizar, mover, marcar lido e ENCAMINHAR (numero + URL do chamado).
Recebem sempre a 'page' do Outlook (nao criam sessao).
"""

import time

from cvc_trata_forms.domain.regras import limpar_titulo
from cvc_trata_forms.infrastructure.sistema.diagnostico import salvar_diagnostico

# rotulo da linha da URL do chamado no corpo do encaminhamento
ROTULO_LINK = "Link chamado: "


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


def _conferir_corpo(corpo, ticket, link):
    """Confere no DOM o que ficou no topo do corpo do encaminhamento.

    {ticket_ok: o numero do chamado esta no texto, url_ok: a URL esta no texto,
    autolink: o OWA transformou a URL em <a> clicavel}. 'autolink' e so
    informativo - a URL em texto puro ja resolve (da p/ copiar/colar e a
    maioria dos leitores de e-mail torna clicavel)."""
    return corpo.evaluate(
        """(el, args) => {
            // normaliza: o innerText quebra linha e o OWA mete zero-width no
            // meio da URL ao auto-linkar -> comparar sem espacos/invisiveis
            const limpa = (s) => (s || '')
                .replace(/[\\u200b-\\u200d\\ufeff\\s]/g, '')
                .replace(/\\/$/, '');
            const texto = limpa(el.innerText);
            const alvoT = limpa(args.ticket);
            const alvoL = limpa(args.link);
            const hrefs = Array.from(el.querySelectorAll('a'))
                               .map(a => limpa(a.getAttribute('href')));
            // o numero do chamado tambem aparece DENTRO da URL
            // (.../browse/GAAR-47) -> tira a URL antes de procurar o numero,
            // senao a linha do numero pode faltar e a conferencia nem ver
            const semUrl = alvoL ? texto.split(alvoL).join('') : texto;
            return {
                ticket_ok: semUrl.includes(alvoT),
                // vale tanto a URL no texto quanto ela no href do auto-link
                // (o OWA as vezes encurta o texto exibido do link)
                url_ok: texto.includes(alvoL) || hrefs.includes(alvoL),
                autolink: hrefs.includes(alvoL),
            };
        }""", {"ticket": ticket, "link": link})


def _compose_aberto(outlook):
    """True enquanto a janela de escrever o e-mail estiver na tela (o campo
    'Para' so existe no compose) - e o sinal de que o envio ainda NAO saiu."""
    return outlook.locator(
        "[contenteditable='true'][aria-label='Para']").count() > 0


def _esperar_compose_fechar(outlook, timeout_s):
    """Espera o compose sumir (= e-mail enviado). True se fechou no prazo."""
    fim = time.time() + timeout_s
    while time.time() < fim:
        if not _compose_aberto(outlook):
            return True
        time.sleep(0.5)
    return False


def _enviar_compose(outlook, log, timeout_s=12):
    """Envia o e-mail e CONFIRMA que saiu (compose fechou).

    Na maquina do cliente o clique no 'Enviar' "da certo" sem enviar nada e o
    e-mail fica em RASCUNHOS - mesmo sintoma que ja tinhamos no envio do
    chamado no Jira. Por isso o sucesso nao e o clique e sim o compose fechar,
    e sao 3 tentativas em escada: clique do Playwright -> Ctrl+Enter (atalho do
    OWA, com o foco no corpo) -> click() disparado no proprio DOM."""
    btn = outlook.get_by_role("button", name="Enviar", exact=False)
    if btn.count() > 0:
        try:
            btn.first.click()
        except Exception as e:
            log(f"   [encaminhar] clique no 'Enviar' falhou ({e}); vou de atalho.")
    else:
        log("   [encaminhar] botao 'Enviar' nao encontrado; vou de atalho.")
    _confirmar_popup_enviar(outlook, log)
    if _esperar_compose_fechar(outlook, timeout_s):
        return True

    # 2a: atalho do teclado. Antes garante o foco DENTRO do compose, senao a
    # tecla nao chega nele (clicar no corpo nao envia nada, e seguro).
    log("   [encaminhar] o clique nao enviou (e-mail ficou em rascunhos); "
        "tentando Ctrl+Enter (atalho do OWA).")
    caixa = outlook.locator("[role='textbox'][aria-label='Corpo da mensagem']")
    if caixa.count() > 0:
        try:
            caixa.first.click()
        except Exception:
            pass
    outlook.keyboard.press("Control+Enter")
    _confirmar_popup_enviar(outlook, log)
    if _esperar_compose_fechar(outlook, timeout_s):
        return True

    # 3a: dispara o click no proprio elemento, pelo DOM (contorna overlay/
    # camada invisivel que engole o clique do mouse)
    log("   [encaminhar] Ctrl+Enter tambem nao enviou; tentando o clique "
        "pelo DOM.")
    if btn.count() > 0:
        try:
            btn.first.evaluate("el => el.click()")
        except Exception as e:
            log(f"   [encaminhar] clique pelo DOM falhou ({e}).")
    _confirmar_popup_enviar(outlook, log)
    return _esperar_compose_fechar(outlook, timeout_s)


# textos que CANCELAM o envio no popup - nunca clicar neles
_NAO_ENVIAR = ("nao enviar", "não enviar", "don't send", "dont send",
               "cancelar", "cancel", "descartar", "discard")


def _confirmar_popup_enviar(outlook, log=None):
    """Trata o popup 'parece que voce esqueceu de anexar' que o OWA mostra ao
    enviar. Clica no botao que CONFIRMA o envio e nunca no que cancela.

    Antes so aceitava o texto EXATO 'Enviar' - se o botao vier como 'Enviar
    mesmo assim' (que e o texto do popup de anexo), o clique nao acontecia e o
    popup ficava na tela segurando o envio. Agora casa por conteudo."""
    time.sleep(1.5)
    dlg = outlook.locator("[role='dialog'], [role='alertdialog']")
    if dlg.count() == 0:
        return False
    botoes = dlg.first.get_by_role("button")
    for i in range(botoes.count()):
        b = botoes.nth(i)
        try:
            if not b.is_visible():
                continue
            txt = (b.inner_text() or "").strip().lower()
        except Exception:
            continue
        if not txt or any(n in txt for n in _NAO_ENVIAR):
            continue
        if "enviar" in txt or "send" in txt:
            try:
                b.click()
                if log:
                    log(f"   [encaminhar] popup de anexo: cliquei em "
                        f"'{txt}'.")
                time.sleep(0.8)
                return True
            except Exception:
                pass
    return False


def _escrever_topo(outlook, corpo, texto):
    """Escreve 'texto' no TOPO do corpo, de uma unica vez.

    Foca o corpo, sobe pro inicio e digita o bloco inteiro numa chamada so -
    sem pausas no meio, que era onde o editor do OWA se re-renderizava (a
    assinatura carrega depois do compose abrir) e comia o primeiro pedaco."""
    corpo.click()
    time.sleep(0.6)          # deixa o editor assentar antes de digitar
    outlook.keyboard.press("Control+Home")
    time.sleep(0.3)
    outlook.keyboard.type(texto)
    time.sleep(0.5)


def encaminhar_email(outlook, indice, destinatarios, assunto, ticket, link,
                     prefixo="Chamado aberto: ",
                     separador="--- Mensagem original (Microsoft Forms) abaixo ---",
                     log=print):
    """Encaminha o e-mail (indice) para 'destinatarios' com o assunto dado e,
    no TOPO do corpo:

        <prefixo> <ticket>
        Link chamado: <link>
        <separador>

    Tudo em TEXTO PURO (o OWA auto-linka a URL sozinho ao dar Enter). Retorna
    True se enviou.

    Compose OWA: Para=[contenteditable][aria-label='Para'];
    Assunto=input[aria-label='Assunto']; Corpo=[role='textbox'][aria-label=
    'Corpo da mensagem']. O 'De' sai automaticamente como a caixa compartilhada.
    Ao enviar pode surgir o popup de anexos -> _confirmar_popup_enviar."""
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

    # Corpo: numero do chamado + URL do chamado, no TOPO
    corpo = outlook.locator(
        "[role='textbox'][aria-label='Corpo da mensagem']").first
    # Texto de UMA vez so. Digitar em blocos (com pausa entre eles) dava tempo
    # do editor do OWA se re-renderizar no meio - a assinatura entra depois de
    # o compose abrir e ENGOLIA o primeiro bloco: o e-mail saia sem a linha do
    # numero, so com a do link. Tudo numa tacada nao deixa essa brecha.
    # (o \n vira Enter; a linha "em branco" leva um espaco porque o OWA
    # colapsa paragrafo vazio.) Sem Ctrl+K: a URL vai por extenso e o OWA
    # auto-linka sozinho.
    # garante 1 espaco entre o prefixo e o numero (o config faz strip)
    pref = (prefixo.rstrip() + " ") if prefixo.strip() else ""
    bloco = f"{pref}{ticket}\n \n{ROTULO_LINK}{link}\n \n"
    if separador:
        bloco += f"{separador}\n \n"
    _escrever_topo(outlook, corpo, bloco)

    # Conferencia do corpo: NAO barra o envio. Sem Ctrl+K nao existe mais o
    # risco de e-mail torto (o link colado na palavra errada); um e-mail com o
    # numero no assunto vale muito mais que e-mail nenhum. Fica o registro.
    conf = _conferir_corpo(corpo, ticket, link)
    if not conf.get("ticket_ok"):
        # o editor engoliu o comeco -> reescreve SO a linha do numero no topo
        log("   [encaminhar] o numero do chamado nao entrou no corpo (editor "
            "ainda estava montando?); reescrevendo a linha.")
        _escrever_topo(outlook, corpo, f"{pref}{ticket}\n \n")
        conf = _conferir_corpo(corpo, ticket, link)
        log(f"   [encaminhar] apos reescrever: numero={conf.get('ticket_ok')}, "
            f"url={conf.get('url_ok')}.")
    if not conf.get("ticket_ok") or not conf.get("url_ok"):
        log(f"   [encaminhar] conferencia do corpo: numero={conf.get('ticket_ok')}, "
            f"url={conf.get('url_ok')}. Envio segue assim mesmo.")
        salvar_diagnostico(outlook, "encaminhar_corpo_incompleto", log)
    elif not conf.get("autolink"):
        # nao e erro: a URL esta la em texto puro, so nao virou <a> no compose
        log("   [encaminhar] o OWA nao auto-linkou a URL; ela vai em texto "
            "puro (continua copiavel/clicavel no leitor do destinatario).")

    # Enviar de verdade: so vale se o compose FECHAR (clique -> Ctrl+Enter)
    if not _enviar_compose(outlook, log):
        log("   [encaminhar] o e-mail NAO saiu: o compose continua aberto "
            "depois do clique e do Ctrl+Enter.")
        salvar_diagnostico(outlook, "encaminhar_nao_enviou", log)
        raise RuntimeError(
            "encaminhamento nao foi enviado (compose continua aberto).")
    log("   [encaminhar] enviado (compose fechou).")
    time.sleep(1.5)
    return True
