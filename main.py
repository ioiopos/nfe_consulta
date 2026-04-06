#!/usr/bin/env python3
"""
NF-e Destinadas - Abbas Tecnologia
Funciona como script Python e como .exe gerado pelo PyInstaller.
"""

import sys
import os
import traceback
from pathlib import Path


def _configurar_paths():
    if getattr(sys, "frozen", False):
        BASE_DIR = Path(sys.executable).parent
        sys.path.insert(0, sys._MEIPASS)
    else:
        BASE_DIR = Path(__file__).parent
        sys.path.insert(0, str(BASE_DIR))

    for pasta in ["data", "data/notas", "logs", "xml", "certs"]:
        (BASE_DIR / pasta).mkdir(parents=True, exist_ok=True)

    os.environ["ABBAS_NFE_DIR"] = str(BASE_DIR)
    return BASE_DIR


if __name__ == "__main__":
    try:
        BASE_DIR = _configurar_paths()

        # Redireciona warnings do PKCS12 — são inofensivos
        import warnings
        warnings.filterwarnings("ignore", category=UserWarning, module="cryptography")

        from src.app import NFEConsultaApp
        app = NFEConsultaApp()
        app.mainloop()

    except Exception as e:
        # Salva o erro em arquivo de log para diagnóstico
        log_path = Path(os.environ.get("ABBAS_NFE_DIR", ".")) / "logs" / "erro_startup.txt"
        log_path.parent.mkdir(parents=True, exist_ok=True)

        erro_completo = traceback.format_exc()
        log_path.write_text(erro_completo, encoding="utf-8")

        # Tenta mostrar erro em janela gráfica
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "Erro ao iniciar — Abbas Tecnologia",
                f"Erro: {e}\n\nDetalhes salvos em:\n{log_path}\n\n"
                f"Verifique se todas as dependências estão instaladas:\n"
                f"  pip install -r requirements.txt"
            )
        except Exception:
            # Último recurso: imprime no console
            print("=" * 60)
            print("ERRO AO INICIAR:")
            print(erro_completo)
            print(f"Log salvo em: {log_path}")
            print("=" * 60)
            input("Pressione Enter para fechar...")
