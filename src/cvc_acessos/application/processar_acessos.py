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
    abrir_inbox_compartilhada, ler_emails,
)
from cvc_acessos.domain.regras import tem_cat_jira, casa_remetente
from cvc_acessos.infrastructure.outlook.acoes import (
    extrair_email, marcar_lido, mover_email, encaminhar_email,
)
from cvc_acessos.infrastructure.jira.chamado import (
    preencher_formulario, enviar_e_capturar_codigo,
)
from cvc_acessos.infrastructure.config.config_app import (
    DRY_RUN, CAIXA, REMETENTE_FILTRO, SO_NAO_LIDOS, SUBPASTA_DESTINO,
    MARCAR_COMO_LIDO, ENCERRAR_SESSOES_NO_FIM, URL_JIRA, JIRA_CAT_PREFIXO,
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


def _plano(n, dados):
    print(f"\n  --- E-MAIL {n} ---")
    print(f"    Resumen     : {dados['titulo']}")
    print(f"    Corpo       : {dados['corpo']}")
    print(f"    Link        : {dados['link'][:70]}...")
    print("    Tipo        : Forms")


def processar(outlook, jira):
    """UM ciclo: abre a caixa, filtra e processa. NAO faz login nem logoff.
    Retorna (encontrados, processados)."""
    aberto = abrir_inbox_compartilhada(outlook, CAIXA)
    print(f">> Caixa de Entrada: {aberto}")

    if DRY_RUN:
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
        from cvc_acessos.application.sessao.encerrar_sessoes import logoff_completo
        print(">> Logoff (Outlook + Jira)...")
        logoff_completo(browser, log=print)


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:  # noqa: BLE001
        print("\n[ERRO] " + str(ex))
        raise SystemExit(1)
