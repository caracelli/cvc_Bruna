"""
Lancador do app (bandeja) com AUTO-INSTALACAO das dependencias.

    python run.py

Na 1a execucao, se faltar alguma biblioteca (pystray, Pillow, playwright,
keyring, uiautomation), ele roda `pip install -r requirements.txt` sozinho
(no MESMO Python) e reinicia. Depois abre normalmente.

Requer internet + pip na 1a vez. (Alternativa sem nada disso: .exe standalone.)
"""
import importlib.util
import os
import subprocess
import sys

RAIZ = os.path.dirname(os.path.abspath(__file__))
REQ = os.path.join(RAIZ, "requirements.txt")

# nome do modulo (import) -> pacote no pip
NECESSARIOS = {
    "pystray": "pystray",
    "PIL": "Pillow",
    "playwright": "playwright",
    "keyring": "keyring",
    "uiautomation": "uiautomation",
}


def _faltando():
    return [mod for mod in NECESSARIOS if importlib.util.find_spec(mod) is None]


def _instalar(faltam):
    print(">> Dependencias faltando:", ", ".join(faltam))
    print(">> Instalando (primeira execucao)... aguarde.\n")
    if os.path.exists(REQ):
        base = [sys.executable, "-m", "pip", "install", "-r", REQ]
    else:
        base = [sys.executable, "-m", "pip", "install", *NECESSARIOS.values()]
    r = subprocess.run(base)
    if r.returncode != 0:
        print("\n>> Tentando instalar no perfil do usuario (--user)...")
        subprocess.run(base + ["--user"])


def main():
    faltam = _faltando()
    # instala 1x e reinicia o processo (pra carregar os pacotes novos)
    if faltam and not os.environ.get("CVC_BOOTSTRAP_DONE"):
        _instalar(faltam)
        os.environ["CVC_BOOTSTRAP_DONE"] = "1"
        os.execv(sys.executable, [sys.executable, os.path.abspath(__file__),
                                  *sys.argv[1:]])

    faltam = _faltando()
    if faltam:
        print("\n[ERRO] Ainda faltam dependencias:", ", ".join(faltam))
        print("Instale manualmente e rode de novo:")
        print(f"   {sys.executable} -m pip install -r requirements.txt")
        sys.exit(1)

    # tudo ok -> abre o app
    sys.path.insert(0, os.path.join(RAIZ, "src"))
    from cvc_acessos.presentation.bandeja import main as app_main
    app_main()


if __name__ == "__main__":
    main()
