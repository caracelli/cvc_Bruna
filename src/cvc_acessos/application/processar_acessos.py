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

USO: python -m cvc_acessos.application.processar_acessos
=============================================================================
"""

import sys
import time

from cvc_acessos.application.sessao.checar_login import checar
from cvc_acessos.infrastructure.credenciais.credenciais import obter_outlook
from cvc_acessos.infrastructure.jira.login_jira import login_jira
from cvc_acessos.infrastructure.outlook.outlook_web import (
    abrir_inbox_compartilhada, abrir_subpasta, ler_emails, pesquisar_forms,
    limpar_pesquisa,
)
from cvc_acessos.domain.regras import tem_cat_jira, casa_remetente
from cvc_acessos.infrastructure.outlook.acoes import (
    extrair_email, marcar_lido, marcar_nao_lido, mover_email, mover_para_inbox,
    encaminhar_email,
)
from cvc_acessos.infrastructure.jira.chamado import (
    preencher_formulario, enviar_e_capturar_codigo, cancelar_chamado,
)
from cvc_acessos.infrastructure.config.config_app import (
    DRY_RUN, DEMO, DEMO_QTD, CAIXA, REMETENTE_FILTRO, SO_NAO_LIDOS,
    SUBPASTA_DESTINO, MARCAR_COMO_LIDO, ENCERRAR_SESSOES_NO_FIM, URL_JIRA,
    JIRA_CAT_PREFIXO, JIRA_PORTAL_BASE, JIRA_TRANSICAO_CANCELAR,
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
    outlook = est["page_outlook"]
    jira = est["page_jira"]
    if not est["outlook"].get("ok"):
        raise RuntimeError(
            "Outlook nao esta logado. Rode o bootstrap de sessoes antes."
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


def processar_demo(outlook, jira):
    """MODO DEMO (self-cleaning): processa ate DEMO_QTD e-mails de VERDADE
    (cria chamado + encaminha + marca lido + move) e DEPOIS DESFAZ (cancela o
    chamado + volta o e-mail ao Inbox nao lido). Demonstra o fluxo completo sem
    deixar residuo. Retorna (feitos, feitos)."""
    print("=" * 60)
    print(f">> MODO DEMO (faz e desfaz) - ate {DEMO_QTD} e-mail(s)")
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
            dados = extrair_email(outlook, e["indice"])
            print(f"\n>> [DEMO] criando chamado p/: {dados['titulo']}")
            jira.bring_to_front()
            jira.goto(URL_JIRA, wait_until="domcontentloaded")   # form de criacao
            time.sleep(2)
            if not preencher_formulario(jira, dados):
                print("   [DEMO] formulario nao preencheu; parando FASE 1.")
                break
            codigo, link = enviar_e_capturar_codigo(jira)
            if codigo == "??-?":
                print("   [DEMO][ERRO] envio do chamado falhou (codigo '??-?'); "
                      "parando FASE 1 p/ nao encaminhar com link/numero errado. "
                      "Veja o diagnostico [enviar]/[form] acima no log.")
                break
            print(f"   [DEMO] chamado criado: {codigo}  ({link})")
            outlook.bring_to_front()
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
            cancelar_chamado(jira, codigo, JIRA_PORTAL_BASE,
                             JIRA_TRANSICAO_CANCELAR, log=print)
        outlook.bring_to_front()
        abrir_subpasta(outlook, CAIXA, SUBPASTA_DESTINO)
        _aguardar_lista(outlook)
        ids = {x[0] for x in feitos}
        for _ in range(len(feitos) + 2):     # move os nossos de volta ao Inbox
            alvo = next((e for e in ler_emails(outlook, 100)
                         if e["id"] in ids), None)
            if not alvo:
                break
            mover_para_inbox(outlook, alvo["indice"])
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
            dados = extrair_email(outlook, e["indice"])
            print(f"\n>> Processando: {dados['titulo']}")
            jira.bring_to_front()
            if not preencher_formulario(jira, dados):
                print("   [ERRO] nao preencheu o formulario; parando.")
                break
            codigo, link = enviar_e_capturar_codigo(jira)
            if codigo == "??-?":
                print("   [ERRO] nao consegui capturar o numero do chamado "
                      "(envio falhou). Abortando este ciclo p/ NAO encaminhar "
                      "com numero/link errado. Veja o diagnostico [enviar]/"
                      "[form] acima no log.")
                break
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
    finally:
        limpar_pesquisa(outlook, CAIXA)


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
        from cvc_acessos.application.sessao.encerrar_sessoes import logoff_completo
        print(">> Logoff (Outlook + Jira)...")
        logoff_completo(browser, log=print)


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:  # noqa: BLE001
        print("\n[ERRO] " + str(ex))
        raise SystemExit(1)
