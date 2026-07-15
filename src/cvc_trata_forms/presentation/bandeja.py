"""
=============================================================================
 APP DE BANDEJA (system tray)  -  CVC-Trata-Forms
=============================================================================
Icone perto do relogio que:
  - faz login (1x, com MFA) ao iniciar;
  - monitora a caixa a cada N minutos (config.xml -> fluxo/intervalo_min);
  - roda o fluxo (DRY-RUN ou real, conforme config.xml -> fluxo/dry_run);
  - menu: Verificar agora / Pausar / Reiniciar / Ver log / Fechar.

"Fechar" faz LOGOFF (encerra Outlook+Jira). Mantem o Edge aberto enquanto
o app estiver rodando (a sessao dura enquanto o Edge vive).

 USO:
     python app_bandeja.py
=============================================================================
"""

import os
import sys
import contextlib
import subprocess
import threading
from datetime import datetime

from PIL import Image, ImageDraw
import pystray


# --- pystray: abrir o menu com o clique ESQUERDO (alem do direito) ---
def _patch_menu_clique_esquerdo():
    try:
        import pystray._win32 as _w
        w = _w.win32

        def _on_notify(self, wparam, lparam):
            if self._menu_handle and lparam in (w.WM_LBUTTONUP,
                                                w.WM_RBUTTONUP):
                w.SetForegroundWindow(self._hwnd)
                point = _w.wintypes.POINT()
                w.GetCursorPos(_w.ctypes.byref(point))
                hmenu, descriptors = self._menu_handle
                index = w.TrackPopupMenuEx(
                    hmenu,
                    w.TPM_RIGHTALIGN | w.TPM_BOTTOMALIGN | w.TPM_RETURNCMD,
                    point.x, point.y, self._menu_hwnd, None)
                if index > 0:
                    descriptors[index - 1](self)

        _w.Icon._on_notify = _on_notify
    except Exception:
        pass


_patch_menu_clique_esquerdo()

import cvc_trata_forms.application.sessao.iniciar_sessoes as iniciar_sessoes
import cvc_trata_forms.application.processar_acessos as fluxo
import cvc_trata_forms.infrastructure.config.config_app as config_app
from cvc_trata_forms.application.sessao.encerrar_sessoes import logoff_completo
from cvc_trata_forms.infrastructure.sistema.winutil import mostrar_edge, esconder_edge
from cvc_trata_forms.infrastructure.sistema.caminhos import RAIZ

LOG_FILE = os.path.join(RAIZ, "cvc_bandeja.log")
INTERVALO_S = max(60, config_app.INTERVALO_MIN * 60)

# ----- estado compartilhado -----
estado = {
    "status": "iniciando...",
    "ultima": "-",
    "encontrados": 0,
    "processados": 0,
    "edge_oculto": config_app.INVISIVEL,
}
pausado = threading.Event()
parar = threading.Event()
disparar = threading.Event()      # forca um ciclo imediato ("Verificar agora")
icone = None


def _abrir_form_config(esperar=False):
    """Abre o formulario de config em processo separado (frozen-aware).
    No .exe, sys.executable e o proprio exe -> usa 'exe --config'.
    esperar=True bloqueia ate o form fechar (usado no setup inicial)."""
    if getattr(sys, "frozen", False):
        cmd = [sys.executable, "--config"]
    else:
        cmd = [sys.executable, "-m", "cvc_trata_forms.presentation.form_config"]
    p = subprocess.Popen(cmd, creationflags=0x08000000)   # CREATE_NO_WINDOW
    if esperar:
        p.wait()
    return p


def log(msg):
    linha = f"{datetime.now():%d/%m %H:%M:%S}  {msg}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(linha + "\n")
    except Exception:
        pass
    try:
        print(linha)          # em --windowed o stdout pode ser None
    except Exception:
        pass


class _LogWriter:
    """file-like: no .exe (--windowed) o stdout vai pro devnull, entao o
    print() do FLUXO se perderia. Grava cada linha DIRETO no LOG_FILE.

    CRITICO: NAO chamar log() aqui. Este writer FICA no lugar do sys.stdout
    durante o fluxo (redirect_stdout); log() faz print(), que voltaria pra ca
    -> recursao infinita (foi o bug do log de 120MB). Escrita direta = sem
    recursao. Guard reentrante por seguranca."""

    def __init__(self):
        self._buf = ""
        self._dentro = False

    def write(self, s):
        if self._dentro:            # evita reentrancia (recursao)
            return
        self._dentro = True
        try:
            self._buf += s
            while "\n" in self._buf:
                linha, self._buf = self._buf.split("\n", 1)
                if linha.strip():
                    try:
                        with open(LOG_FILE, "a", encoding="utf-8") as f:
                            f.write(f"{datetime.now():%d/%m %H:%M:%S}  "
                                    f"{linha.rstrip()}\n")
                    except Exception:
                        pass
        finally:
            self._dentro = False

    def flush(self):
        pass


# ----- icone -----
def _img(cor):
    img = Image.new("RGB", (64, 64), (255, 255, 255))
    d = ImageDraw.Draw(img)
    d.ellipse((8, 8, 56, 56), fill=cor)
    d.text((20, 22), "CVC", fill=(255, 255, 255))
    return img


def _cor_status():
    if estado["status"].startswith("erro") or "falha" in estado["status"]:
        return (200, 60, 60)        # vermelho
    if pausado.is_set():
        return (150, 150, 150)      # cinza
    if "verificando" in estado["status"] or "login" in estado["status"]:
        return (230, 170, 40)       # amarelo
    return (0, 141, 166)            # teal (monitorando)


def atualizar_icone():
    if icone is None:
        return
    modo = ("DEMO" if config_app.DEMO
            else "DRY-RUN" if config_app.DRY_RUN else "REAL")
    icone.title = (f"CVC-Trata-Forms [{modo}]\n"
                   f"Status: {estado['status']}\n"
                   f"Ultima: {estado['ultima']} | "
                   f"enc:{estado['encontrados']} proc:{estado['processados']}")
    try:
        icone.icon = _img(_cor_status())
        icone.update_menu()
    except Exception:
        pass


# ----- login e ciclo -----
def fazer_login():
    estado["status"] = "login (faca o MFA)..."
    atualizar_icone()
    log("Iniciando login - aprove o MFA na janela do Edge.")
    try:
        iniciar_sessoes.main()
        estado["status"] = "monitorando"
        log("Login concluido.")
        atualizar_icone()
        return True
    except Exception as e:  # noqa: BLE001
        estado["status"] = "falha no login"
        log(f"Falha no login: {e}")
        atualizar_icone()
        return False


def ciclo():
    """Roda um ciclo. Retorna True se OK, False se falhou (sessao/Edge)."""
    estado["status"] = "verificando..."
    atualizar_icone()
    pw = None
    ok = False
    try:
        # captura o print() interno do fluxo -> LOG_FILE (no .exe stdout=devnull)
        with contextlib.redirect_stdout(_LogWriter()):
            browser, outlook, jira = fluxo._garantir_sessoes()
            pw = getattr(browser, "_pw_session", None)
            enc, proc = fluxo.processar(outlook, jira)
        estado["ultima"] = datetime.now().strftime("%H:%M")
        estado["encontrados"] = enc
        estado["processados"] = proc
        estado["status"] = "monitorando"
        log(f"Ciclo OK: {enc} encontrado(s), {proc} processado(s).")
        ok = True
    except Exception as e:  # noqa: BLE001
        estado["status"] = "erro (ver log)"
        log(f"Ciclo com erro: {e}")
    finally:
        if pw is not None:
            try:
                pw.stop()
            except Exception:
                pass
    atualizar_icone()
    return ok


def loop_monitor():
    # Login SEMPRE do zero ao iniciar: limpa qualquer sessao anterior (mesmo
    # que o app anterior tenha sido morto/travado sem logoff). Regra
    # "nunca deixar conectado".
    estado["status"] = "limpando sessao anterior..."
    atualizar_icone()
    log("Limpando sessao anterior (login sempre do zero ao abrir)...")
    try:
        logoff_completo(log=log)
    except Exception as e:  # noqa: BLE001
        log(f"(limpeza inicial: {e})")

    fazer_login()
    while not parar.is_set():
        if not pausado.is_set():
            ok = ciclo()
            if not ok and not parar.is_set():
                # Edge fechou / sessao caiu -> reabre o Edge e re-loga sozinho.
                # NAO re-roda o ciclo imediatamente: espera o intervalo (ou um
                # "Verificar agora"). Motivo: um erro NO MEIO do ciclo (ex.: a
                # FASE 2 do DEMO falhando ao mover o e-mail de volta) tambem cai
                # aqui; re-rodar na hora criaria chamados/encaminhamentos
                # duplicados em loop.
                log("Ciclo falhou (Edge fechado/sessao caiu) - "
                    "reabrindo o Edge e re-logando...")
                estado["status"] = "recuperando (re-login)..."
                atualizar_icone()
                if icone is not None:
                    try:
                        icone.notify("Sessao caiu. Reabrindo o Edge - "
                                     "aprove o MFA se pedir.",
                                     "CVC-Trata-Forms")
                    except Exception:
                        pass
                fazer_login()
        # espera o intervalo OU um disparo manual ("Verificar agora")
        disparar.wait(timeout=INTERVALO_S)
        disparar.clear()


# ----- acoes do menu -----
def on_verificar(icon, item):
    log("Verificacao manual solicitada.")
    disparar.set()


def on_toggle_edge(icon, item):
    if estado.get("edge_oculto", True):
        mostrar_edge()
        estado["edge_oculto"] = False
        log("Navegador MOSTRADO.")
    else:
        esconder_edge()
        estado["edge_oculto"] = True
        log("Navegador ESCONDIDO.")
    atualizar_icone()


def on_pausar(icon, item):
    if pausado.is_set():
        pausado.clear()
        log("Monitor RETOMADO.")
    else:
        pausado.set()
        log("Monitor PAUSADO.")
    atualizar_icone()


def on_reiniciar(icon, item):
    log("Reiniciar: refazendo login...")
    threading.Thread(target=fazer_login, daemon=True).start()


def on_config(icon, item):
    """Abre o formulario de configuracao (processo separado, nao trava o app).
    Alteracoes salvas valem no PROXIMO inicio do app (config lido na abertura)."""
    log("Abrindo Configuracoes...")
    try:
        _abrir_form_config()
    except Exception as e:
        log(f"Nao consegui abrir o form: {e}")


def on_log(icon, item):
    try:
        os.startfile(LOG_FILE)
    except Exception as e:
        log(f"Nao consegui abrir o log: {e}")


def on_sair(icon, item):
    log("Encerrando: logoff e fechando...")
    parar.set()
    disparar.set()
    try:
        logoff_completo(log=log)
    except Exception as e:
        log(f"Erro no logoff: {e}")
    icon.stop()


def _menu():
    return pystray.Menu(
        pystray.MenuItem(lambda i: f"Status: {estado['status']}", None,
                         enabled=False),
        pystray.MenuItem(lambda i: f"Ultima: {estado['ultima']}  "
                                   f"(enc {estado['encontrados']} / "
                                   f"proc {estado['processados']})", None,
                         enabled=False),
        pystray.MenuItem(lambda i: ("Modo: DEMO (faz e desfaz)" if config_app.DEMO
                                    else "Modo: DRY-RUN" if config_app.DRY_RUN
                                    else "Modo: REAL"), None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Verificar agora", on_verificar),
        pystray.MenuItem(
            lambda i: ("Mostrar navegador" if estado.get("edge_oculto", True)
                       else "Esconder navegador"),
            on_toggle_edge),
        pystray.MenuItem("Pausar", on_pausar,
                         checked=lambda i: pausado.is_set()),
        pystray.MenuItem("Reiniciar (re-login)", on_reiniciar),
        pystray.MenuItem("Configurações...", on_config),
        pystray.MenuItem("Ver log", on_log),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Fechar", on_sair),
    )


def _config_faltando():
    """Retorna a lista de campos essenciais do config.xml que estao vazios."""
    faltam = []
    if not config_app.OUTLOOK_EMAIL.strip():
        faltam.append("E-mail do Outlook")
    if not config_app.OUTLOOK_SENHA.strip():
        faltam.append("Senha")
    if not config_app.URL_OUTLOOK.strip():
        faltam.append("URL do Outlook")
    if not config_app.URL_JIRA.strip():
        faltam.append("URL do Jira")
    if not config_app.CAIXA.strip():
        faltam.append("Caixa compartilhada")
    if config_app.ENCAMINHAR_ATIVO and not config_app.ENCAMINHAR_DESTINATARIOS:
        faltam.append("Destinatários do encaminhamento")
    return faltam


def _definir_app_id():
    """Windows: define um AppUserModelID explicito p/ as notificacoes/janelas
    aparecerem com o NOME DO APP em vez de 'Python' (que e o processo hospedeiro
    quando roda pela fonte). No .exe empacotado ja aparece o nome do exe, mas
    isso deixa consistente tambem rodando pela fonte. So-Windows; nunca lanca."""
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "CVC.TrataForms")
    except Exception:
        pass


def main():
    global icone
    _definir_app_id()
    # 1a execucao / config incompleta -> abre o FORM direto (nao inicia o monitor)
    faltam = _config_faltando()
    if faltam:
        log("Config incompleta: " + ", ".join(faltam) + " -> abrindo formulario.")
        try:
            import tkinter as tk
            from tkinter import messagebox
            r = tk.Tk()
            r.withdraw()
            messagebox.showinfo(
                "Configuração inicial",
                "Faltam dados:\n\n  - " + "\n  - ".join(faltam) +
                "\n\nVou abrir o formulário. Preencha e clique Salvar — o "
                "aplicativo inicia sozinho em seguida.")
            r.destroy()
        except Exception:
            pass
        try:
            _abrir_form_config(esperar=True)      # espera o form fechar
        except Exception as e:  # noqa: BLE001
            log(f"Nao consegui abrir o form: {e}")
        # recarrega o config do disco e reavalia
        try:
            import importlib
            importlib.reload(config_app)
        except Exception:
            pass
        if not _config_faltando():
            # ok agora -> reinicia numa instancia NOVA (recarrega tudo)
            log("Config salva -> reiniciando o app automaticamente.")
            try:
                subprocess.Popen([sys.executable] + sys.argv[1:],
                                 creationflags=0x08000000)
            except Exception as e:  # noqa: BLE001
                log(f"Nao consegui reiniciar: {e}")
        else:
            try:
                import tkinter as tk
                from tkinter import messagebox
                r = tk.Tk()
                r.withdraw()
                messagebox.showinfo(
                    "Configuração",
                    "Ainda faltam dados. Reabra o aplicativo quando preencher.")
                r.destroy()
            except Exception:
                pass
        return

    log(f"App iniciado. Intervalo {config_app.INTERVALO_MIN} min | "
        f"{'DEMO' if config_app.DEMO else 'DRY-RUN' if config_app.DRY_RUN else 'REAL'}")
    icone = pystray.Icon("cvc_trata_forms", _img(_cor_status()),
                         "CVC-Trata-Forms", _menu())
    threading.Thread(target=loop_monitor, daemon=True).start()
    icone.run()


if __name__ == "__main__":
    main()
