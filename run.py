"""
Atalho para rodar SEM instalar (adiciona src/ ao path e abre a bandeja).
    python run.py
(Para instalar de verdade: pip install -e .  e depois use  cvc-acessos)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from cvc_acessos.presentation.bandeja import main  # noqa: E402

if __name__ == "__main__":
    main()
