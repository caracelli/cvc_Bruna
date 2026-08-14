"""
=============================================================================
 CASO DE USO — Processar Acessos (CVC / Gestao de Acessos)
=============================================================================
Orquestra o fluxo, apoiando-se nas camadas de baixo:
  - dominio      -> regras puras (filtro / titulo / categoria)
  - infra/outlook-> ler caixa + acoes (extrair, encaminhar, mover, lido)
  - infra/jira   -> preencher + enviar chamado

Para cada e-mail "Microsoft Forms" NAO LIDO na caixa compartilhada:
  1. extrai titulo/corpo/link
  2. preenche e ENVIA o chamado no Jira -> codigo + link
  3. ENCAMINHA o e-mail pro grupo (nº do chamado no assunto + link clicavel)
  4. marca como lido e move para a subpasta "Finalizados"

DRY_RUN=True: so LE + PREVIEW do formulario (nao envia/encaminha/move).

USO: python -m cvc_trata_forms.application.processar_acessos
=============================================================================
"""

import sys
import time

from cvc_trata_forms.application.sessao.checar_login import checar
from cvc_trata_forms.infrastructure.credenciais.credenciais import obter_outlook
from cvc_trata_forms.infrastructure.jira.login_jira import login_jira
from cvc_trata_forms.infrastructure.outlook.outlook_web import (
    abrir_inbox_compartilhada, abrir_subpasta, ler_emails, pesquisar_forms,
    limpar_pesquisa,
)
from cvc_trata_forms.domain.regras import tem_cat_jira, casa_remetente
from cvc_trata_forms.infrastructure.outlook.acoes import (
    extrair_email, marcar_lido, marcar_nao_lido, mover_email, mover_para_inbox,
    encaminhar_email,
)
from cvc_trata_forms.infrastructure.jira.chamado import (
    preencher_formulario, enviar_e_capturar_codigo, cancelar_chamado,
)
from cvc_trata_forms.infrastructure.jira import chamado_api
from cvc_trata_forms.infrastructure.config.config_app import (
    DRY_RUN, DEMO, DEMO_QTD, CAIXA, REMETENTE_FILTRO, SO_NAO_LIDOS,
    SUBPASTA_DESTINO, MARCAR_COMO_LIDO, ENCERRAR_SESSOES_NO_FIM, URL_JIRA,
    JIRA_CAT_PREFIXO, JIRA_PORTAL_BASE, JIRA_TRANSICAO_CANCELAR, JIRA_API_ATIVA,
    ENCAMINHAR_ATIVO, ENCAMINHAR_DESTINATARIOS, ENCAMINHAR_ASSUNTO,
    ENCAMINHAR_PREFIXO, ENCAMINHAR_SEPARADOR,
)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


# ---------------------------------------------------------------------------
#  SESSAO
# ---------------------------------------------------------------------------
def _garantir_sessoes():
    est = checar()
    browser = est["browser"]
    # O OWA renderiza a arvore/navegacao de forma ASSINCRONA apos o login; a 1a
    # leitura (logo apos "Login concluido") pode ler "nao logado" numa sessao
    # FRIA. Nao declarar falha de cara - re-checa por ~25s REUSANDO o browser
    # (sem abrir nova sessao Playwright) ate o Outlook responder.
    fim = time.time() + 25
    while not est["outlook"].get("ok") and time.time() < fim:
        time.sleep(1.5)
        est = checar(browser)
    outlook = est["page_outlook"]
    jira = est["page_jira"]
    if not est["outlook"].get("ok"):
        raise RuntimeError(
            "Outlook nao esta logado. Rode o bootstrap de sessoes antes."
        )
    # Com a API do Jira ativa, NAO precisamos da aba/login do Jira.
    if not JIRA_API_ATIVA and not est["jira"].get("ok"):
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
    return tem_cat_jira(email.get("categorias", []), JIRA_CAT_PREFIXO)


def _filtrar(outlook):
    return [
        e for e in ler_emails(outlook, 100)
        if casa_remetente(e["resumo"], REMETENTE_FILTRO)
        and (not SO_NAO_LIDOS or e["nao_lido"])
        and not _tem_cat_jira(e)
    ]


def _aguardar_lista(outlook):
    """Espera a lista da caixa compartilhada carregar (assincrona)."""
    for _ in range(15):
        if ler_emails(outlook, 5):
            return
        time.sleep(2)


def _focar(page):
    """Traz a aba (Outlook/Jira) para a frente no Edge, para acompanhar o fluxo
    ao vivo: Outlook enquanto le/encaminha o e-mail, Jira enquanto cria o
    chamado. Silencioso e tolerante a falha (nao interrompe o fluxo)."""
    try:
        page.bring_to_front()
        time.sleep(0.3)
    except Exception:
        pass


def _criar_chamado(jira, dados):
    """Cria o chamado e retorna (codigo, link). Usa a API do Jira quando
    configurada (JIRA_API_ATIVA) — sem abrir a aba/formulario — ou cai no
    fluxo de navegador. Retorna ('??-?', '') em caso de falha."""
    if JIRA_API_ATIVA:
        return chamado_api.criar_chamado(dados)
    # --- fallback: formulario no navegador ---
    _focar(jira)
    jira.goto(URL_JIRA, wait_until="domcontentloaded")
    time.sleep(2)
    if not preencher_formulario(jira, dados):
        return "??-?", ""
    return enviar_e_capturar_codigo(jira)


def _cancelar_chamado(jira, codigo):
    """Cancela o chamado (usado no DEMO). Via API quando ativa, senao navegador."""
    if JIRA_API_ATIVA:
        return chamado_api.cancelar_chamado(codigo)
    _focar(jira)
    return cancelar_chamado(jira, codigo, JIRA_PORTAL_BASE,
                            JIRA_TRANSICAO_CANCELAR, log=print)


def processar_demo(outlook, jira):
    """MODO DEMO (self-cleaning): processa ate DEMO_QTD e-mails de VERDADE
    (cria chamado + encaminha + marca lido + move) e DEPOIS DESFAZ (cancela o
    chamado + volta o e-mail ao Inbox nao lido). Demonstra o fluxo completo sem
    deixar residuo. Retorna (feitos, feitos)."""
    print("=" * 60)
    print(f">> MODO DEMO (faz e desfaz) - ate {DEMO_QTD} e-mail(s)")
    _focar(outlook)                  # mostra o Outlook enquanto busca o e-mail
    abrir_inbox_compartilhada(outlook, CAIXA)
    _aguardar_lista(outlook)
    n = pesquisar_forms(outlook, REMETENTE_FILTRO or "Microsoft Forms")
    print(f">> Pesquisa '{REMETENTE_FILTRO}' (pasta atual): {n} resultado(s)")

    # ---- FASE 1: processa de VERDADE ----
    feitos = []                      # lista de (id, resumo, codigo)
    try:
        while len(feitos) < DEMO_QTD:
            ja = {x[0] for x in feitos}
            alvo = [e for e in _filtrar(outlook) if e["id"] not in ja]
            if not alvo:
                break
            e = alvo[0]
            _focar(outlook)          # mostra o Outlook enquanto le o e-mail
            dados = extrair_email(outlook, e["indice"])
            print(f"\n>> [DEMO] criando chamado p/: {dados['titulo']}")
            codigo, link = _criar_chamado(jira, dados)
            if codigo == "??-?":
                print("   [DEMO][ERRO] criacao do chamado falhou (codigo '??-?'); "
                      "parando FASE 1 p/ nao encaminhar com link/numero errado.")
                break
            print(f"   [DEMO] chamado criado: {codigo}  ({link})")
            _focar(outlook)          # volta ao Outlook p/ encaminhar/mover
            if ENCAMINHAR_ATIVO and ENCAMINHAR_DESTINATARIOS:
                assunto = ENCAMINHAR_ASSUNTO.format(ticket=codigo,
                                                    assunto=dados["titulo"])
                encaminhar_email(outlook, e["indice"], ENCAMINHAR_DESTINATARIOS,
                                 assunto, codigo, link, ENCAMINHAR_PREFIXO,
                                 ENCAMINHAR_SEPARADOR)
                print(f"   [DEMO] encaminhado p/ {ENCAMINHAR_DESTINATARIOS}")
            marcar_lido(outlook, e["indice"])
            mover_email(outlook, e["indice"], SUBPASTA_DESTINO)
            print(f"   [DEMO] movido p/ '{SUBPASTA_DESTINO}' + marcado lido")
            feitos.append((e["id"], e["resumo"], codigo))
    finally:
        limpar_pesquisa(outlook, CAIXA)
    print(f"\n>> [DEMO] FASE 1: {len(feitos)} processado(s): "
          f"{[c for _, _, c in feitos]}")

    # ---- FASE 2: DESFAZ (cancela chamado + volta e-mail nao lido) ----
    if feitos:
        print(">> [DEMO] FASE 2: cancelando chamados e voltando os e-mails...")
        for _id, _res, codigo in feitos:
            _cancelar_chamado(jira, codigo)
        _focar(outlook)              # volta ao Outlook p/ restaurar o e-mail
        abrir_subpasta(outlook, CAIXA, SUBPASTA_DESTINO)
        _aguardar_lista(outlook)
        ids = {x[0] for x in feitos}
        for _ in range(len(feitos) + 2):     # move os nossos de volta ao Inbox
            alvo = next((e for e in ler_emails(outlook, 100)
                         if e["id"] in ids), None)
            if not alvo:
                break
            if not mover_para_inbox(outlook, alvo["indice"]):
                # nao conseguiu mover este e-mail de volta -> tira da fila p/
                # NAO ficar tentando o mesmo em loop (ja salvou diagnostico).
                print(f"   [DEMO][AVISO] nao movi de volta o e-mail id={alvo['id']}; "
                      "veja a pasta 'diagnostico'. Segue com os demais.")
                ids.discard(alvo["id"])
            time.sleep(1)
        abrir_inbox_compartilhada(outlook, CAIXA)
        _aguardar_lista(outlook)
        for e in ler_emails(outlook, 100):   # restaura NAO lido (so os nossos)
            if e["id"] in ids and not e["nao_lido"]:
                marcar_nao_lido(outlook, e["indice"])
        print(f">> [DEMO] concluido: {len(feitos)} ciclo(s) feito(s) e DESFEITO(s). "
              "Caixa e Jira de volta ao estado original.")
    return len(feitos), len(feitos)


def processar(outlook, jira):
    """UM ciclo: abre a caixa, filtra e processa. NAO faz login nem logoff.
    Retorna (encontrados, processados)."""
    if DEMO:
        return processar_demo(outlook, jira)
    _focar(outlook)                  # mostra o Outlook enquanto busca o e-mail
    aberto = abrir_inbox_compartilhada(outlook, CAIXA)
    print(f">> Caixa de Entrada: {aberto}")

    # A caixa COMPARTILHADA carrega de forma ASSINCRONA (a lista aparece uns
    # segundos depois de abrir). Sem esperar, _filtrar leria 0. Espera ~30s.
    for _ in range(15):
        if ler_emails(outlook, 5):
            break
        time.sleep(2)

    # Isola a fila de trabalho: PESQUISA 'Microsoft Forms' com escopo 'Pasta
    # atual'. A lista passa a ter SO os Forms do Inbox (sem outros remetentes,
    # sem os processados de 'Finalizados', sem virtualizacao). O nao-lido
    # continua no _filtrar (SO_NAO_LIDOS). limpar_pesquisa restaura a caixa.
    n = pesquisar_forms(outlook, REMETENTE_FILTRO or "Microsoft Forms")
    print(f">> Pesquisa '{REMETENTE_FILTRO}' (pasta atual): {n} resultado(s)")

    try:
        if DRY_RUN:
            # DRY-RUN 100% NAO-INVASIVO: NAO abre os e-mails (abrir marca como
            # lido no painel de leitura). So lista, da propria linha, o que
            # SERIA feito. Sem tocar no Jira e sem alterar a caixa.
            alvo = _filtrar(outlook)
            print(f">> {len(alvo)} e-mail(s) Microsoft Forms NAO lidos "
                  "que SERIAM processados:\n")
            for n, e in enumerate(alvo, 1):
                print(f"  {n}. {e['resumo'][:70]}")
                print("     [DRY-RUN] criaria o chamado no Jira + pegaria o codigo")
                if ENCAMINHAR_ATIVO and ENCAMINHAR_DESTINATARIOS:
                    print(f"     [DRY-RUN] encaminharia p/ {ENCAMINHAR_DESTINATARIOS} "
                          "(nº do chamado no assunto + link clicavel)")
                print(f"     [DRY-RUN] marcaria lido + moveria p/ '{SUBPASTA_DESTINO}'")
            print("\n>> DRY-RUN: nada aberto/enviado/movido. A caixa NAO foi alterada.")
            return len(alvo), 0

        # EXECUCAO REAL: processa o 1o da fila ate acabar (mover tira da lista)
        encontrados = len(_filtrar(outlook))
        processados = 0
        while True:
            # a cada volta re-le os resultados da pesquisa: ao marcar lido +
            # mover, o e-mail deixa de ser 'Forms nao lido na pasta atual' e sai
            # da fila -> nunca reprocessa e nao perde nenhum.
            alvo = _filtrar(outlook)
            if not alvo:
                break
            e = alvo[0]
            _focar(outlook)          # mostra o Outlook enquanto le o e-mail
            dados = extrair_email(outlook, e["indice"])
            print(f"\n>> Processando: {dados['titulo']}")
            codigo, link = _criar_chamado(jira, dados)
            if codigo == "??-?":
                print("   [ERRO] nao consegui criar o chamado. Abortando este "
                      "ciclo p/ NAO encaminhar com numero/link errado.")
                break
            print(f"   Chamado criado: {codigo}  ({link})")
            _focar(outlook)          # volta ao Outlook p/ encaminhar/mover
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
            if not JIRA_API_ATIVA and jira is not None:
                jira.goto(URL_JIRA, wait_until="domcontentloaded")
                time.sleep(2)

        print(f"\n>> {processados} e-mail(s) processado(s).")
        return encontrados, processados
    finally:
        limpar_pesquisa(outlook, CAIXA)


def main():
    modo = "[ DRY-RUN / ENSAIO ]" if DRY_RUN else "[ EXECUCAO REAL ]"
    print("=" * 74)
    print(f"  FLUXO CVC / GESTAO DE ACESSOS   {modo}")
    print("=" * 74)

    print(">> Verificando sessoes...")
    browser, outlook, jira = _garantir_sessoes()
    if JIRA_API_ATIVA:
        print("   OK: Outlook logado. Jira via API (sem aba/login).\n")
    else:
        print("   OK: Outlook e Jira logados.\n")

    processar(outlook, jira)

    if not DRY_RUN and ENCERRAR_SESSOES_NO_FIM:
        from cvc_trata_forms.application.sessao.encerrar_sessoes import logoff_completo
        print(">> Logoff (Outlook + Jira)...")
        logoff_completo(browser, log=print)


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:  # noqa: BLE001
        print("\n[ERRO] " + str(ex))
        raise SystemExit(1)
