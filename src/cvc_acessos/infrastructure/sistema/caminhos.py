"""
Caminhos PORTAVEIS do projeto (rodar em qualquer maquina do cliente sem
interferir no ambiente dele).

- Detecta o Edge automaticamente (varios locais possiveis + PATH).
- Usa um PERFIL dedicado em %LOCALAPPDATA% (isolado do Edge pessoal do
  usuario; nao mexe no perfil/dados do navegador dele).
- Nada de caminho fixo de maquina especifica.
"""

import os
import shutil
import sys
from pathlib import Path

PORTA = 9222


def _achar_raiz():
    """Onde ficam os dados locais (config.xml, credenciais.xml, log).
    - .exe (PyInstaller/frozen): a PROPRIA pasta do executavel.
    - codigo: a pasta que contem config.example.xml / pyproject.toml."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    p = Path(__file__).resolve()
    for cand in [p, *p.parents]:
        if (cand / "config.example.xml").exists() \
                or (cand / "pyproject.toml").exists():
            return cand
    return Path.cwd()


RAIZ = str(_achar_raiz())

# Pasta dedicada (isolada) para o perfil do Edge de automacao.
# Fica em %LOCALAPPDATA%\cvc_gestao_acessos\edge_profile (gravavel, por usuario).
_BASE = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
PERFIL = os.path.join(_BASE, "cvc_gestao_acessos", "edge_profile")

# Marcador unico para identificar (e so afetar) os processos DESTE perfil.
MARCADOR_PERFIL = "cvc_gestao_acessos"


def _edge_locais_conhecidos():
    return [
        os.path.join(os.environ.get("ProgramFiles(x86)", ""),
                     "Microsoft", "Edge", "Application", "msedge.exe"),
        os.path.join(os.environ.get("ProgramFiles", ""),
                     "Microsoft", "Edge", "Application", "msedge.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""),
                     "Microsoft", "Edge", "Application", "msedge.exe"),
    ]


def _edge_via_registro():
    """Le o caminho do Edge no Registro (App Paths) - jeito mais confiavel."""
    try:
        import winreg
    except Exception:
        return None
    chaves = [
        (winreg.HKEY_LOCAL_MACHINE,
         r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"),
        (winreg.HKEY_LOCAL_MACHINE,
         r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"),
        (winreg.HKEY_CURRENT_USER,
         r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"),
    ]
    for hive, sub in chaves:
        try:
            with winreg.OpenKey(hive, sub) as k:
                val, _ = winreg.QueryValueEx(k, None)
                if val and os.path.exists(val):
                    return val
        except OSError:
            pass
    return None


def _edge_via_varredura():
    """Ultimo recurso: varre as pastas de programas atras do msedge.exe."""
    raizes = [
        os.environ.get("ProgramFiles", ""),
        os.environ.get("ProgramFiles(x86)", ""),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft"),
    ]
    for raiz in raizes:
        if not raiz or not os.path.isdir(raiz):
            continue
        for dirpath, _dirs, arquivos in os.walk(raiz):
            if "msedge.exe" in arquivos:
                return os.path.join(dirpath, "msedge.exe")
    return None


def achar_edge():
    """Caminho do msedge.exe (ou None). Ordem: locais conhecidos -> PATH ->
    Registro -> varredura no disco (ultimo recurso)."""
    for c in _edge_locais_conhecidos():
        if c and os.path.exists(c):
            return c
    achado = shutil.which("msedge")
    if achado:
        return achado
    via_reg = _edge_via_registro()
    if via_reg:
        return via_reg
    return _edge_via_varredura()


EDGE = achar_edge()
