"""
Fecha baloes da INTERFACE do Edge (chrome do navegador) que o Playwright nao
alcanca: "Entrar no Microsoft Edge", sync, etc. Usa UI Automation do Windows
e clica SO em botoes de dispensar, em VARIOS IDIOMAS (a maquina do cliente
pode estar em PT, EN ou ES).

Seguro: so procura dentro de janelas do Edge (classe Chrome_WidgetWin*) e so
clica em botoes cujo nome esta na lista de "dispensar". Nao fecha abas/janelas
(nao usamos textos genericos tipo "Fechar"/"Close").
"""

# Textos de botoes que DISPENSAM o balao (multi-idioma). Adicionar conforme
# necessario. NAO inclua "Fechar/Close/Cerrar" (poderia fechar aba/janela).
DISPENSAR = [
    # Portugues
    "Não, obrigado", "Nao, obrigado", "Agora não", "Agora nao",
    # Ingles
    "No thanks", "No, thanks", "Not now", "Maybe later",
    # Espanhol
    "No, gracias", "No gracias", "Ahora no", "Ahora, no", "No por ahora",
    "Quizás más tarde", "Quizas mas tarde",
]

# AutomationIds estaveis (independem de idioma) - preencher se descobertos.
DISPENSAR_IDS = ["noThanks", "no-thanks", "declineButton", "secondaryButton"]


def dispensar_popups_edge(log=lambda *_: None):
    """Procura e clica botoes de dispensar baloes do Edge. Retorna quantos."""
    try:
        import uiautomation as auto
    except Exception:
        return 0

    fechados = 0
    try:
        for w in auto.GetRootControl().GetChildren():
            try:
                cls = w.ClassName or ""
            except Exception:
                cls = ""
            if "Chrome_WidgetWin" not in cls:   # so janelas do Edge/Chromium
                continue

            clicou = False
            # 1) por AutomationId (independe de idioma)
            for aid in DISPENSAR_IDS:
                try:
                    btn = w.ButtonControl(searchDepth=25, AutomationId=aid)
                    if btn.Exists(0, 0):
                        btn.Click(simulateMove=False, waitTime=0)
                        fechados += 1
                        clicou = True
                        log(f"   popup fechado (id={aid})")
                        break
                except Exception:
                    pass
            if clicou:
                continue

            # 2) por texto do botao (multi-idioma)
            for alvo in DISPENSAR:
                try:
                    btn = w.ButtonControl(searchDepth=25, Name=alvo)
                    if btn.Exists(0, 0):
                        btn.Click(simulateMove=False, waitTime=0)
                        fechados += 1
                        log(f"   popup fechado ('{alvo}')")
                        break
                except Exception:
                    pass
    except Exception:
        pass
    return fechados


if __name__ == "__main__":
    import sys
    import time
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    if "--loop" in sys.argv:
        # roda continuamente em PROCESSO PROPRIO (nao bloqueia o login).
        while True:
            try:
                dispensar_popups_edge()
            except Exception:
                pass
            time.sleep(3)
    else:
        n = dispensar_popups_edge(log=print)
        print(f">> {n} popup(s) fechado(s).")
