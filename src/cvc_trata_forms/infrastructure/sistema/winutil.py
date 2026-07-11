"""
Utilitarios de janela (Windows) para o Edge de AUTOMACAO.

Acha SOMENTE as janelas do Edge do perfil de automacao (pelo PID dos processos
cujo command line tem o MARCADOR_PERFIL) e permite:
  - mostrar_edge()  -> exibe e traz para frente (mesmo se estava escondida)
  - esconder_edge() -> esconde totalmente (invisivel, fora da barra de tarefas)
  - trazer_edge_frente() -> alias de mostrar_edge()

Seguro: nunca toca em outras janelas do usuario.
"""

import os
import subprocess

from cvc_trata_forms.infrastructure.sistema.caminhos import MARCADOR_PERFIL

SW_HIDE = 0
SW_SHOW = 5
SW_RESTORE = 9


def _pids_edge_automacao():
    try:
        ps = ("Get-CimInstance Win32_Process -Filter \"Name='msedge.exe'\" | "
              f"Where-Object {{ $_.CommandLine -like '*{MARCADOR_PERFIL}*' }} | "
              "Select-Object -ExpandProperty ProcessId")
        out = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                             capture_output=True, text=True, timeout=10,
                             creationflags=0x08000000)   # CREATE_NO_WINDOW
        return set(int(x) for x in out.stdout.split() if x.strip().isdigit())
    except Exception:
        return set()


def _janelas_edge(user32):
    """HWNDs das janelas do Edge de automacao (pela CLASSE, nao pelo titulo,
    pois logo apos abrir a janela ainda nao tem titulo). Acha visiveis E
    escondidas."""
    import ctypes
    pids = _pids_edge_automacao()
    achados = []
    if not pids:
        return achados

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def _cb(hwnd, _):
        pid = ctypes.c_ulong(0)
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value not in pids:
            return True
        buf = ctypes.create_unicode_buffer(64)
        user32.GetClassNameW(hwnd, buf, 64)
        if buf.value == "Chrome_WidgetWin_1":   # janela principal do Edge
            achados.append(hwnd)
        return True

    user32.EnumWindows(_cb, 0)
    return achados


def mostrar_edge():
    """Mostra e traz a janela do Edge de automacao para frente."""
    if os.name != "nt":
        return
    try:
        import ctypes
        user32 = ctypes.windll.user32
        janelas = _janelas_edge(user32)
        if not janelas:
            return
        user32.SetWindowPos.argtypes = [
            ctypes.c_void_p, ctypes.c_void_p, ctypes.c_int, ctypes.c_int,
            ctypes.c_int, ctypes.c_int, ctypes.c_uint,
        ]
        SWP = 0x0001 | 0x0002 | 0x0040     # NOSIZE | NOMOVE | SHOWWINDOW
        TOPMOST = ctypes.c_void_p(-1)
        NOTOPMOST = ctypes.c_void_p(-2)
        for hwnd in janelas:
            user32.ShowWindow(hwnd, SW_SHOW)        # DESFAZ o SW_HIDE
            user32.ShowWindow(hwnd, SW_RESTORE)     # restaura se minimizada
            user32.SetWindowPos(hwnd, TOPMOST, 0, 0, 0, 0, SWP)
            user32.SetWindowPos(hwnd, NOTOPMOST, 0, 0, 0, 0, SWP)
            user32.SetForegroundWindow(hwnd)
            user32.BringWindowToTop(hwnd)
    except Exception:
        pass


def esconder_edge():
    """Esconde TOTALMENTE a janela do Edge de automacao (invisivel).
    Re-tenta por ~6s pois logo apos abrir a janela pode ainda nao existir."""
    if os.name != "nt":
        return
    import time
    try:
        import ctypes
        user32 = ctypes.windll.user32
        for _ in range(12):
            janelas = _janelas_edge(user32)
            if janelas:
                for hwnd in janelas:
                    user32.ShowWindow(hwnd, SW_HIDE)
                return
            time.sleep(0.5)
    except Exception:
        pass


def trazer_edge_frente():
    """Compatibilidade: mostra e traz o Edge para frente."""
    mostrar_edge()
