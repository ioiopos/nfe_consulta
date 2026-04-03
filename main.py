#!/usr/bin/env python3
"""
NF-e Destinadas - Abbas Tecnologia
Funciona como script Python e como .exe gerado pelo PyInstaller.
"""

import sys
import os
from pathlib import Path


def _configurar_paths():
    if getattr(sys, "frozen", False):
        # .exe PyInstaller — BASE ao lado do executavel
        BASE_DIR = Path(sys.executable).parent
        sys.path.insert(0, sys._MEIPASS)
    else:
        # Script normal
        BASE_DIR = Path(__file__).parent
        sys.path.insert(0, str(BASE_DIR))

    # Cria pastas de dados ao lado do exe/script
    for pasta in ["data", "data/notas", "logs", "xml", "certs"]:
        (BASE_DIR / pasta).mkdir(parents=True, exist_ok=True)

    os.environ["ABBAS_NFE_DIR"] = str(BASE_DIR)


_configurar_paths()

from src.app import NFEConsultaApp

if __name__ == "__main__":
    app = NFEConsultaApp()
    app.mainloop()
