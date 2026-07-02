"""
Modulo de automacao do Outlook web (PLANO B) via Edge/Chrome em depuracao.

Reaproveita a sua sessao ja logada (incl. caixas compartilhadas), sem API,
sem admin e sem consentimento. Conecta no navegador aberto por
'abrir_edge.ps1' (ou 'abrir_chrome.ps1') na porta 9222.

Funcoes principais:
    conectar()                      -> retorna (browser, page) do Outlook
    listar_pastas(page)             -> lista itens da navegacao (pastas/caixas)
    abrir_pasta(page, nome)         -> clica na pasta/caixa pelo nome
    ler_emails(page, limite)        -> retorna lista de dicts com os e-mails

Uso tipico:
    from cvc_acessos.infrastructure.outlook.outlook_web import conectar, abrir_pasta, ler_emails
    browser, page = conectar()
    abrir_pasta(page, "Gestão de Acessos")
    for e in ler_emails(page, 20):
        print(e["resumo"])
"""

import re
import sys
import time

from playwright.sync_api import sync_playwright

# Garante que acentos/icones nao quebrem a saida no console do Windows
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

PORTA_CDP = "http://127.0.0.1:9222"


def conectar():
    """Conecta no navegador em depuracao e devolve (browser, page_outlook).

    Mantemos o objeto 'playwright' vivo dentro do browser para a sessao
    nao ser fechada. Use 'browser' apenas para nao perder a referencia.
    """
    p = sync_playwright().start()
    try:
        browser = p.chromium.connect_over_cdp(PORTA_CDP)
    except Exception:
        raise RuntimeError(
            "Nao consegui conectar na porta 9222. Abra o navegador com "
            "'abrir_edge.ps1' (ou abrir_chrome.ps1) e faca login no Outlook."
        )

    page = None
    for ctx in browser.contexts:
        for pg in ctx.pages:
            if "outlook" in pg.url or "office.com" in pg.url:
                page = pg
                break
    if page is None:
        raise RuntimeError(
            "Conectei no navegador, mas nao achei a aba do Outlook. "
            "Abra https://outlook.office.com/mail/ e faca login."
        )

    page.bring_to_front()
    # guarda a sessao playwright junto pra nao ser coletada
    browser._pw_session = p  # type: ignore[attr-defined]
    return browser, page


def listar_pastas(page):
    """Retorna a lista de rotulos dos itens de navegacao (pastas e caixas)."""
    tree = page.locator("div[role='treeitem']")
    itens = []
    for i in range(tree.count()):
        lbl = (tree.nth(i).get_attribute("aria-label")
               or tree.nth(i).inner_text() or "").strip().replace("\n", " ")
        itens.append(lbl)
    return itens


def abrir_pasta(page, nome, timeout_ms=8000):
    """Clica na pasta/caixa cujo rotulo CONTEM 'nome' (case-insensitive).

    Retorna o rotulo completo do item aberto. Lanca erro se nao achar.
    """
    alvo = nome.lower()
    # A caixa COMPARTILHADA carrega de forma assincrona apos o login, entao
    # re-tenta por ~25s antes de desistir.
    fim = time.time() + 25
    while time.time() < fim:
        tree = page.locator("div[role='treeitem']")
        for i in range(tree.count()):
            lbl = (tree.nth(i).get_attribute("aria-label")
                   or tree.nth(i).inner_text() or "")
            if alvo in lbl.lower():
                tree.nth(i).click()
                _esperar_lista(page, timeout_ms)
                return lbl.strip().replace("\n", " ")
        # nao achou: expande os nos colapsados (caixa compartilhada / Inbox)
        # pra revelar subpastas como 'Finalizados' e tenta de novo
        _expandir_colapsados(page)
        time.sleep(1.5)
    raise RuntimeError(
        f"Nao encontrei a pasta/caixa contendo '{nome}'. "
        f"Pastas disponiveis: {listar_pastas(page)}"
    )


def _arvore_labels(page):
    """Lista (indice, rotulo) de todos os itens da arvore de pastas."""
    tree = page.locator("div[role='treeitem']")
    out = []
    for i in range(tree.count()):
        lbl = (tree.nth(i).get_attribute("aria-label")
               or tree.nth(i).inner_text() or "").strip().replace("\n", " ")
        out.append((i, lbl))
    return tree, out


def abrir_inbox_compartilhada(page, nome_caixa, timeout_ms=8000):
    """Abre a CAIXA DE ENTRADA REAL da caixa compartilhada.

    ATENCAO: o no com o nome da caixa (ex.: 'Gestao de Acessos - Argentina')
    e apenas o CABECALHO da conta — clicar nele NAO troca a lista de
    mensagens. A caixa de entrada de verdade e a subpasta logo abaixo
    ('Inbox' / 'Caixa de Entrada'). Esta funcao mira essa subpasta.

    Retorna o rotulo da pasta aberta. Lanca erro se nao achar.
    """
    alvo = nome_caixa.lower()
    inbox_termos = ("inbox", "caixa de entrada", "entrada")
    fim = time.time() + 25
    while time.time() < fim:
        _expandir_colapsados(page)
        tree, labels = _arvore_labels(page)
        idx_hdr = next((i for i, lbl in labels if alvo in lbl.lower()), None)
        if idx_hdr is not None:
            # 1a subpasta de entrada DEPOIS do cabecalho da caixa
            for i, lbl in labels:
                if i > idx_hdr and any(t in lbl.lower() for t in inbox_termos):
                    tree.nth(i).click()
                    _esperar_lista(page, timeout_ms)
                    return lbl
        time.sleep(1.5)
    raise RuntimeError(
        f"Nao achei a Caixa de Entrada da caixa '{nome_caixa}'. "
        f"Pastas: {listar_pastas(page)}"
    )


def abrir_subpasta(page, nome_caixa, subpasta, timeout_ms=8000):
    """Abre uma SUBPASTA especifica da caixa compartilhada (ex.: 'Finalizados'),
    procurando-a SOMENTE depois do cabecalho da caixa (evita casar com pastas
    pessoais de mesmo nome). Retorna o rotulo aberto."""
    alvo = nome_caixa.lower()
    sub = subpasta.lower()
    fim = time.time() + 25
    while time.time() < fim:
        _expandir_colapsados(page)
        tree, labels = _arvore_labels(page)
        idx_hdr = next((i for i, lbl in labels if alvo in lbl.lower()), None)
        ini = (idx_hdr + 1) if idx_hdr is not None else 0
        for i, lbl in labels:
            if i >= ini and sub in lbl.lower():
                tree.nth(i).click()
                _esperar_lista(page, timeout_ms)
                return lbl
        time.sleep(1.5)
    raise RuntimeError(
        f"Nao achei a subpasta '{subpasta}' da caixa '{nome_caixa}'. "
        f"Pastas: {listar_pastas(page)}"
    )


def _expandir_colapsados(page):
    """Expande nos da arvore que estao recolhidos (aria-expanded='false'),
    revelando subpastas (ex.: 'Finalizados' dentro da caixa compartilhada).
    Usa foco + seta direita pra expandir SEM navegar/perder a selecao."""
    tree = page.locator("div[role='treeitem']")
    mudou = False
    for i in range(tree.count()):
        try:
            no = tree.nth(i)
            if no.get_attribute("aria-expanded") == "false":
                no.focus()
                page.keyboard.press("ArrowRight")
                time.sleep(0.5)
                mudou = True
        except Exception:
            pass
    return mudou


def _esperar_lista(page, timeout_ms=8000):
    """Espera a lista de mensagens recarregar apos trocar de pasta."""
    fim = time.time() + timeout_ms / 1000
    ultimo = -1
    estavel = 0
    while time.time() < fim:
        n = page.locator("div[role='option']").count()
        if n == ultimo:
            estavel += 1
            if estavel >= 2:  # contagem estavel por ~0.6s
                return
        else:
            estavel = 0
            ultimo = n
        time.sleep(0.3)


def _parse_aria(aria):
    """Quebra o aria-label de uma linha de e-mail em campos uteis.

    O formato do Outlook costuma ser:
      [Tem anexos] <Remetente> <Assunto> <Data/Hora> <Previa...>
    Como nao ha separadores fixos, devolvemos o texto bruto (resumo) e
    tentamos extrair remetente (primeiro trecho) de forma best-effort.
    """
    aria = (aria or "").strip()
    tem_anexo = aria.lower().startswith("tem anexos")
    return {
        "tem_anexo": tem_anexo,
        "resumo": aria,
    }


def _esta_nao_lido(item, aria):
    """Detecta se a linha de e-mail esta NAO LIDA (robusto, multi-idioma).

    Regra principal: o botao de alternancia mostra a ACAO OPOSTA ao estado.
      - "Marcar como lido" / "Mark as read"   -> mensagem NAO LIDA
      - "Marcar como nao lido" / "Mark as unread" -> mensagem LIDA
    Fallback: aria-label da linha comeca com "Nao lid"/"Unread".
    """
    if item.locator(
        "[title='Marcar como lido'], [title='Mark as read']"
    ).count() > 0:
        return True
    a = (aria or "").lower()
    return a.startswith("não lid") or a.startswith("nao lid") or a.startswith("unread")


# titulo do "chip" de categoria na linha (multi-idioma)
_TITULO_CAT = ("[title*='com a categoria' i], [title*='with the category' i], "
               "[title*='con la categoría' i], [title*='con la categoria' i]")
_RE_CAT = re.compile(r"(?:categoria|category|categoría)\s+(.+)$", re.I)


def categorias_do_email(item):
    """Le as categorias aplicadas no e-mail RAPIDO (sem abrir menu).
    Cada categoria aparece como um chip com title tipo
    'Pesquisar todas as mensagens com a categoria <NOME>'."""
    cats = []
    try:
        chips = item.locator(_TITULO_CAT)
        for i in range(chips.count()):
            t = chips.nth(i).get_attribute("title") or ""
            m = _RE_CAT.search(t)
            if m:
                cats.append(m.group(1).strip())
    except Exception:
        pass
    return cats


def tem_categoria(item, codigo):
    """True se o e-mail tem a categoria 'codigo' (rapido, via title da linha)."""
    try:
        return item.locator(f"[title*='{codigo}' i]").count() > 0
    except Exception:
        return False


def ler_emails(page, limite=20):
    """Retorna ate 'limite' e-mails da pasta atualmente aberta.

    Cada item: {indice, id, tem_anexo, nao_lido, categorias, resumo}.
    """
    opts = page.locator("div[role='option']")
    n = min(limite, opts.count())
    emails = []
    for i in range(n):
        item = opts.nth(i)
        aria = item.get_attribute("aria-label") or ""
        dados = _parse_aria(aria)
        dados["indice"] = i
        dados["id"] = item.get_attribute("data-convid") or ""
        dados["nao_lido"] = _esta_nao_lido(item, aria)
        dados["categorias"] = categorias_do_email(item)   # rapido, sem menu
        emails.append(dados)
    return emails
