"""
Entrypoint UNICO do .exe (PyInstaller). Um so executavel, 3 modos por argumento:

    cvc-acessos.exe                 -> abre o app de bandeja (padrao)
    cvc-acessos.exe --config        -> abre o formulario de configuracao
    cvc-acessos.exe --fechar-popups -> loop do fechador de popups do Edge

(No modo .exe, sys.executable e o proprio .exe, entao os subprocessos internos
chamam 'sys.executable --config' / '--fechar-popups' em vez de 'python -m ...'.)
"""
import os
import sys
import time

# rodando do codigo (nao congelado): garante src/ no path
if not getattr(sys, "frozen", False):
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    "src"))


def main():
    args = sys.argv[1:]
    if "--check" in args:
        # verifica que TODAS as libs pesadas foram empacotadas (nao roda nada)
        import pystray  # noqa: F401
        import PIL  # noqa: F401
        import keyring  # noqa: F401
        import uiautomation  # noqa: F401
        from playwright.sync_api import sync_playwright  # noqa: F401
        import cvc_acessos.presentation.bandeja  # noqa: F401
        print("CHECK OK: pystray, PIL, keyring, uiautomation, playwright, "
              "cvc_acessos - tudo empacotado.")
        return
    if "--fechar-popups" in args:
        from cvc_acessos.infrastructure.sistema.fechar_popups import (
            dispensar_popups_edge,
        )
        while True:
            try:
                dispensar_popups_edge()
            except Exception:
                pass
            time.sleep(3)
    elif "--config" in args:
        from cvc_acessos.presentation.form_config import main as form_main
        form_main()
    else:
        from cvc_acessos.presentation.bandeja import main as bandeja_main
        bandeja_main()


if __name__ == "__main__":
    main()
