"""
Infra / Jira — preenchimento e envio do chamado no portal Service Desk.
Recebe sempre a 'page' do Jira (nao cria sessao/login).
"""

import re
import time

from cvc_trata_forms.infrastructure.config.config_app import (
    JIRA_TIPO_SOLICITUD, JIRA_COR_TITULO, JIRA_PORTAL_BASE, JIRA_PREFIXO_CHAMADO,
)
from cvc_trata_forms.infrastructure.sistema.diagnostico import salvar_diagnostico


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


def preencher_formulario(jira, dados, log=print):
    """Preenche Resumen + Descripcion (cabecalho teal + corpo + link) + Tipo.
    NAO envia. Retorna True se preencheu."""
    _fechar_cookies(jira)
    if jira.locator("#summary").count() == 0:
        log("   [form] campo #summary nao encontrado (form nao carregou?).")
        salvar_diagnostico(jira, "form_nao_carregou", log)
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

    # Tipo de solicitud (CAMPO OBRIGATORIO: se nao selecionar, o botao Enviar
    # fica DESABILITADO e o envio falha silenciosamente -> codigo '??-?').
    if not _selecionar_tipo(jira, log):
        log("   [form] AVISO: 'Tipo de solicitud' NAO foi selecionado; o botao "
            "Enviar pode ficar DESABILITADO e o envio falhar.")
    return True


def _selecionar_tipo(jira, log=print):
    """Seleciona o 'Tipo de solicitud' (obrigatorio). Tenta pelo id conhecido
    (fragil) e, se falhar, por rotulo (pt/es/en). Retorna True se marcou uma
    opcao. Digita JIRA_TIPO_SOLICITUD e escolhe a opcao que contem 'form'
    (ou a 1a opcao, como ultimo recurso)."""
    campos = [
        jira.locator("#pf-undefined-cd-26"),
        jira.get_by_label("Tipo de solicitud", exact=False),
        jira.get_by_label("Tipo de solicitação", exact=False),
        jira.get_by_label("Tipo de solicitacao", exact=False),
        jira.get_by_label("Request type", exact=False),
    ]
    for campo in campos:
        try:
            if campo.count() == 0:
                continue
            campo.first.click()
            time.sleep(0.6)
            jira.keyboard.type(JIRA_TIPO_SOLICITUD)
            time.sleep(1.6)
            ops = jira.locator("[role='option']")
            for i in range(ops.count()):
                txt = (ops.nth(i).inner_text() or "")
                if "form" in txt.lower():
                    ops.nth(i).click()
                    log(f"   [form] tipo selecionado: '{txt.strip()[:40]}'")
                    return True
            if ops.count() > 0:
                txt = ops.first.inner_text() or ""
                ops.first.click()
                log(f"   [form] tipo (1a opcao): '{txt.strip()[:40]}'")
                return True
        except Exception:
            continue
    return False


def _status_chamado(jira):
    """Le o badge de status do chamado (texto curto). '?' se nao achar."""
    for sel in ["[data-test-id*='status' i]", "[class*='status' i] span"]:
        loc = jira.locator(sel)
        for i in range(min(loc.count(), 6)):
            t = (loc.nth(i).inner_text() or "").strip()
            if t and len(t) < 40:
                return t
    return "?"


def cancelar_chamado(jira, codigo, portal_base, transicao="Cancelado pelo Solicitante",
                     log=print):
    """CANCELA o chamado 'codigo' pelo portal (transicao do solicitante).
    Usado no modo DEMO p/ desfazer. VALIDADO ao vivo (GAAR-11..22):
      1. abre o chamado (portal_base + codigo)
      2. clica a transicao p/ ABRIR o popup — clique INSTAVEL, com RETRY (5x)
         ate o popup aparecer (placeholder 'Comentário opcional' OU 2+ botoes
         de mesmo rotulo)
      3. CONFIRMA no ULTIMO botao visivel de rotulo=transicao (o overlay
         renderiza por ultimo; o botao 'Cancelar' do popup NAO cancela o chamado)
      4. verifica pelo BADGE de status (reload). Retorna True se cancelou.
    """
    jira.goto(portal_base + codigo, wait_until="domcontentloaded")
    time.sleep(6)
    if jira.get_by_role("button", name=transicao, exact=False).count() == 0:
        return "cancel" in _status_chamado(jira).lower()
    aberto = False
    for _ in range(5):
        try:
            jira.get_by_role("button", name=transicao, exact=False).first.click()
        except Exception:
            pass
        for _ in range(8):
            if (jira.get_by_placeholder("Comentário opcional").count() > 0
                    or jira.get_by_role("button", name=transicao,
                                        exact=False).count() >= 2):
                aberto = True
                break
            time.sleep(0.5)
        if aberto:
            break
        time.sleep(1)
    cands = jira.get_by_role("button", name=transicao, exact=False)
    for i in reversed(range(cands.count())):
        try:
            if cands.nth(i).is_visible():
                cands.nth(i).click()
                break
        except Exception:
            pass
    time.sleep(4)
    jira.reload(wait_until="domcontentloaded")
    time.sleep(4)
    ok = "cancel" in _status_chamado(jira).lower()
    log(f"   Cancelamento {codigo}: {'OK' if ok else 'nao confirmado'}")
    return ok


def _listar_botoes(jira, limite=40):
    """(texto, visivel, habilitado) dos buttons da pagina — p/ diagnostico."""
    out = []
    b = jira.locator("button")
    for i in range(min(b.count(), limite)):
        try:
            txt = (b.nth(i).inner_text() or "").strip()
            if txt:
                out.append((txt[:35], b.nth(i).is_visible(),
                            b.nth(i).is_enabled()))
        except Exception:
            pass
    return out


def _enviar_por_enter(jira, log=print):
    """Fallback de submit por TECLADO: em algumas maquinas/portais o botao
    Enviar nao responde ao clique programatico, mas o formulario E enviado ao
    pressionar ENTER (confirmado ao vivo na maquina do cliente). Tenta o ENTER
    com o foco atual (onde a selecao do 'Tipo' deixou) e depois focando o campo
    Resumen (#summary, input de 1 linha -> Enter aciona o submit implicito).
    NUNCA usa o editor de descricao (Enter la so quebra linha).
    Retorna True (best-effort; quem confirma o envio e a captura do codigo)."""
    try:
        log("   [enviar] fallback: enviando via tecla ENTER (foco atual).")
        jira.keyboard.press("Enter")
        time.sleep(1.5)
    except Exception as e:
        log(f"   [enviar] ENTER (foco atual) falhou: {e}")
    try:
        s = jira.locator("#summary")
        if s.count() > 0:
            s.first.click()
            time.sleep(0.3)
            log("   [enviar] fallback: ENTER com foco no campo Resumen.")
            jira.keyboard.press("Enter")
            time.sleep(1.0)
    except Exception as e:
        log(f"   [enviar] ENTER (Resumen) falhou: {e}")
    return True


def _clicar_enviar(jira, log=print):
    """Envia o formulario. 1o tenta CLICAR o botao (VISIVEL + HABILITADO, rola
    ate ele, retry pois a validacao do form e assincrona; por rotulo pt/es/en
    e por ultimo button[type=submit]). Se nenhum clique disparar, cai no
    fallback por ENTER (_enviar_por_enter). Retorna True se conseguiu enviar
    (por clique OU por Enter); loga os botoes disponiveis antes do fallback."""
    rotulos = ["Enviar", "Criar", "Crear", "Create", "Send"]
    for _ in range(4):
        for nome in rotulos:
            cand = jira.get_by_role("button", name=nome, exact=False)
            for i in range(cand.count()):
                b = cand.nth(i)
                try:
                    if b.is_visible() and b.is_enabled():
                        b.scroll_into_view_if_needed(timeout=2000)
                        b.click()
                        log(f"   [enviar] cliquei no botao '{nome}'.")
                        return True
                except Exception:
                    continue
        sub = jira.locator("button[type='submit']")
        for i in range(sub.count()):
            b = sub.nth(i)
            try:
                if b.is_visible() and b.is_enabled():
                    b.scroll_into_view_if_needed(timeout=2000)
                    b.click()
                    log("   [enviar] cliquei via button[type=submit].")
                    return True
            except Exception:
                continue
        time.sleep(1.5)     # aguarda a validacao assincrona habilitar o botao
    log("   [enviar] NENHUM botao de envio respondeu ao clique. Botoes na pagina:")
    for txt, vis, en in _listar_botoes(jira):
        log(f"      - '{txt}' visivel={vis} habilitado={en}")
    # fallback por teclado: o clique nao dispara em algumas maquinas, mas o
    # ENTER envia o formulario (comportamento confirmado na maquina do cliente).
    return _enviar_por_enter(jira, log)


def enviar_e_capturar_codigo(jira, log=print, timeout_s=40):
    """Clica em Enviar e captura (codigo, link) do chamado criado.

    IMPORTANTE (validado ao vivo, GAAR-27): ESTE portal do JSM NAO redireciona
    apos o envio — a URL continua no formulario (.../create/<n>) e o chamado
    criado aparece como CONFIRMACAO no CORPO da pagina (ex.: 'GAAR-27'). Por
    isso NAO da pra capturar pela URL nem derivar o link dela.

    Estrategia:
      1. clica Enviar de forma robusta (visivel+habilitado, com retry).
      2. AGUARDA o codigo surgir — pela URL (caso algum portal redirecione) OU
         pelo CORPO — o que vier primeiro (poll a cada 0.5s). Assim capturamos
         assim que o chamado e criado, sem esperar um redirect que nao vem.
      3. deriva o LINK do chamado como JIRA_PORTAL_BASE + codigo (mesma forma
         usada — e validada — por cancelar_chamado).

    Retorna (codigo, link). Se o envio/captura falhar, retorna ('??-?', url_atual)
    e LOGA o diagnostico (o caller aborta p/ nao encaminhar com dados errados)."""
    url_antes = jira.url
    # Se um prefixo de chamado esta configurado (ex.: 'GAAR'), a captura SO
    # aceita tokens com esse prefixo -> blinda contra pegar um 'ABC-123' que
    # ja estava no corpo do e-mail Forms. Vazio = aceita qualquer prefixo.
    pref = (JIRA_PREFIXO_CHAMADO or "").strip().rstrip("-")
    corpo_pat = re.escape(pref) + r"-\d+" if pref else r"[A-Z]{2,8}-\d+"
    padrao_url = re.compile(r"/portal/\d+/(" + corpo_pat + r")")
    padrao_txt = re.compile(r"\b" + corpo_pat + r"\b")
    log(f"   [enviar] URL antes do envio: {url_antes}")

    if not _clicar_enviar(jira, log):
        log("   [enviar] FALHA: formulario NAO foi enviado (botao Enviar "
            "indisponivel; provavel campo obrigatorio faltando). Codigo='??-?'.")
        salvar_diagnostico(jira, "botao_enviar_indisponivel", log)
        return "??-?", jira.url

    fim = time.time() + timeout_s
    while time.time() < fim:
        # 1) algum portal redireciona p/ a pagina do chamado -> pega da URL
        m = padrao_url.search(jira.url)
        if m and jira.url != url_antes:
            link = jira.url.split("?")[0]
            log(f"   [enviar] chamado criado (via URL): {m.group(1)}  ({link})")
            return m.group(1), link
        # 2) ESTE portal: o codigo aparece no CORPO (banner de confirmacao)
        try:
            txt = jira.locator("body").inner_text()
        except Exception:
            txt = ""
        m2 = padrao_txt.search(txt)
        if m2:
            codigo = m2.group(0)
            link = JIRA_PORTAL_BASE + codigo
            log(f"   [enviar] chamado criado (via corpo): {codigo}  ({link})")
            return codigo, link
        time.sleep(0.5)

    log(f"   [enviar] codigo NAO encontrado apos {timeout_s}s (URL: {jira.url}) "
        "-> '??-?'.")
    salvar_diagnostico(jira, "codigo_nao_capturado", log)
    return "??-?", jira.url
