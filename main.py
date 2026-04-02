#!/usr/bin/env python3
"""
NF-e Destinadas - Consulta SEFAZ
Ponto de entrada da aplicação.

Uso:
    python main.py
"""

import sys
import os

# Garante que o diretório do projeto está no path
sys.path.insert(0, os.path.dirname(__file__))

from src.app import NFEConsultaApp

if __name__ == "__main__":
    app = NFEConsultaApp()
    app.mainloop()
