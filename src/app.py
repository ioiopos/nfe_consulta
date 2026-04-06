"""
NF-e Destinadas — Abbas Tecnologia
Interface em CustomTkinter com design Claymorphism real.
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import json
import os
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from src.certificado import CertificadoDigital
from src.sefaz_client import SefazClient
from src.storage import StorageNSU
from src.evento_client import EventoClient, TIPOS_EVENTO
from src.win_cert_store import listar_certificados_windows, IS_WINDOWS

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

ACCENT     = "#e63428"
ACCENT2    = "#c0281d"
ACCENT_LT  = "#fff0ee"
DARK       = "#1a1a2e"
BG         = "#f5f0ff"
CARD       = "#ffffff"
INPUT_BG   = "#f8f9ff"
SUCCESS    = "#1a7a4a"
SUCCESS_LT = "#e8fff4"
WARNING    = "#d48806"
ERROR      = "#c0392b"
ERROR_LT   = "#fff0f0"
TEXT       = "#1a1a2e"
TEXT_DIM   = "#6b7a99"
BORDER     = "#e8ecf4"
LOG_BG     = "#1a1a2e"

FONT_TITLE = ("Segoe UI", 14, "bold")
FONT_HEAD  = ("Segoe UI", 10, "bold")
FONT_BODY  = ("Segoe UI", 10)
FONT_SMALL = ("Segoe UI", 9)
FONT_MONO  = ("Consolas", 9)
FONT_BADGE = ("Segoe UI", 9, "bold")


class NFEConsultaApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Abbas Tecnologia · NF-e Destinadas — SEFAZ Ambiente Nacional")
        self.geometry("1280x800")
        self.minsize(1060, 700)
        self.configure(fg_color=BG)
        self.cert    = None
        self.storage = StorageNSU()
        self._build_ui()
        self._atualizar_status_cert()
        self._carregar_nfes_salvas()
        self._restaurar_cert_salvo()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_header()

        body = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        body.pack(fill="both", expand=True, padx=18, pady=14)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        left = ctk.CTkScrollableFrame(body, width=300, fg_color=BG,
                                       corner_radius=0,
                                       scrollbar_button_color=BORDER,
                                       scrollbar_button_hover_color=ACCENT_LT)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 14))

        right = ctk.CTkFrame(body, fg_color=BG, corner_radius=0)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)

        self._build_panel_cert(left)
        self._build_panel_consulta(left)
        self._build_panel_nsu(left)
        self._build_panel_resultados(right)
        self._build_panel_log(right)
        self._build_status_bar()

    def _build_header(self):
        hdr = ctk.CTkFrame(self, fg_color=DARK, corner_radius=0, height=64)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        inner = ctk.CTkFrame(hdr, fg_color="transparent", corner_radius=0)
        inner.pack(fill="x", padx=22, pady=10)

        logo = ctk.CTkFrame(inner, fg_color=ACCENT, corner_radius=14, width=44, height=44)
        logo.pack(side="left", padx=(0, 14))
        logo.pack_propagate(False)
        ctk.CTkLabel(logo, text="A", font=("Segoe UI", 20, "bold"), text_color="white").pack(expand=True)

        tf = ctk.CTkFrame(inner, fg_color="transparent", corner_radius=0)
        tf.pack(side="left")
        ctk.CTkLabel(tf, text="Abbas Tecnologia", font=("Segoe UI", 14, "bold"), text_color="white").pack(anchor="w")
        ctk.CTkLabel(tf, text="NF-e Destinadas  ·  Modelo 55  ·  SEFAZ Ambiente Nacional",
                     font=("Segoe UI", 8), text_color="#8899cc").pack(anchor="w")

        badge = ctk.CTkFrame(inner, fg_color=ACCENT, corner_radius=20)
        badge.pack(side="right")
        ctk.CTkLabel(badge, text="  Consulta Fiscal  ", font=FONT_BADGE, text_color="white").pack(padx=4, pady=6)

        ctk.CTkFrame(self, fg_color=ACCENT, height=4, corner_radius=0).pack(fill="x")

    def _card(self, parent, titulo):
        frame = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=16,
                              border_width=2, border_color="#eef0f8")
        frame.pack(fill="x", pady=(0, 12), padx=2)
        pill = ctk.CTkFrame(frame, fg_color=ACCENT_LT, corner_radius=20)
        pill.pack(anchor="w", padx=14, pady=(12, 6))
        ctk.CTkLabel(pill, text=f"  {titulo}  ", font=FONT_BADGE, text_color=ACCENT).pack(padx=4, pady=4)
        inner = ctk.CTkFrame(frame, fg_color="transparent", corner_radius=0)
        inner.pack(fill="both", expand=True, padx=14, pady=(0, 12))
        return inner

    def _lbl(self, parent, text):
        ctk.CTkLabel(parent, text=text, font=FONT_SMALL, text_color=TEXT_DIM).pack(anchor="w", pady=(4, 1))

    def _entry(self, parent, textvariable=None, show=None, width=None):
        kw = dict(textvariable=textvariable, fg_color=INPUT_BG, border_color=BORDER,
                  border_width=2, text_color=TEXT, corner_radius=10, font=FONT_BODY)
        if show:  kw["show"] = show
        if width: kw["width"] = width
        e = ctk.CTkEntry(parent, **kw)
        e.pack(fill="x", pady=(2, 8))
        return e

    def _btn(self, parent, text, command, fg=None, fg2=None, tc="white", h=38):
        b = ctk.CTkButton(parent, text=text, command=command,
                          fg_color=fg or ACCENT, hover_color=fg2 or ACCENT2,
                          text_color=tc, corner_radius=12, height=h,
                          font=FONT_HEAD)
        b.pack(fill="x", pady=(0, 8))
        return b

    def _btn_ghost(self, parent, text, command):
        b = ctk.CTkButton(parent, text=text, command=command,
                          fg_color=ACCENT_LT, hover_color="#ffe0cc",
                          text_color=ACCENT, border_color=ACCENT,
                          border_width=2, corner_radius=10, height=34, font=FONT_SMALL)
        b.pack(fill="x", pady=(0, 8))
        return b

    # ── Painéis ───────────────────────────────────────────────────────────────

    def _build_panel_cert(self, parent):
        card = self._card(parent, "CERTIFICADO DIGITAL")
        self._lbl(card, "Arquivo (.pfx / .p12):")
        frow = ctk.CTkFrame(card, fg_color="transparent", corner_radius=0)
        frow.pack(fill="x", pady=(2, 8))
        self.var_cert_path = tk.StringVar(value="Nenhum certificado carregado")
        ctk.CTkLabel(frow, textvariable=self.var_cert_path, font=FONT_SMALL,
                     text_color=TEXT_DIM, wraplength=210, justify="left").pack(side="left", fill="x", expand=True)
        ctk.CTkButton(frow, text="📁", width=36, height=32, fg_color=INPUT_BG,
                      hover_color=ACCENT_LT, text_color=ACCENT, corner_radius=8,
                      font=("Segoe UI", 14), command=self._selecionar_cert).pack(side="right")

        self._lbl(card, "Senha do certificado:")
        self.var_senha = tk.StringVar()
        e = self._entry(card, textvariable=self.var_senha, show="●")
        e.bind("<Return>", lambda _: self._carregar_cert())

        self._btn(card, "Carregar certificado", self._carregar_cert)
        ctk.CTkFrame(card, fg_color=BORDER, height=1, corner_radius=0).pack(fill="x", pady=(4, 10))

        if IS_WINDOWS:
            self._btn_ghost(card, "Usar certificado instalado no Windows",
                            self._selecionar_cert_windows)

        self.lbl_cert_status = ctk.CTkLabel(card, text="Nenhum certificado carregado",
                                             font=FONT_SMALL, text_color=TEXT_DIM)
        self.lbl_cert_status.pack(anchor="w", pady=(4, 0))

    def _build_panel_consulta(self, parent):
        card = self._card(parent, "CONSULTA SEFAZ")
        self._lbl(card, "CNPJ (somente números):")
        self.var_cnpj = tk.StringVar()
        self.entry_cnpj = self._entry(card, textvariable=self.var_cnpj)

        uf_row = ctk.CTkFrame(card, fg_color="transparent", corner_radius=0)
        uf_row.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(uf_row, text="Cód. UF (IBGE):", font=FONT_SMALL, text_color=TEXT_DIM).pack(side="left")
        self.var_uf = tk.StringVar(value="43")
        ctk.CTkEntry(uf_row, textvariable=self.var_uf, width=52, fg_color=INPUT_BG,
                     border_color=BORDER, border_width=2, text_color=TEXT,
                     corner_radius=8, font=FONT_BODY, height=30).pack(side="left", padx=(6, 8))
        ctk.CTkLabel(uf_row, text="(43=RS 35=SP 41=PR...)", font=FONT_SMALL, text_color=TEXT_DIM).pack(side="left")

        self._lbl(card, "Ambiente:")
        self.var_ambiente = tk.StringVar(value="1")
        amb = ctk.CTkFrame(card, fg_color="transparent", corner_radius=0)
        amb.pack(anchor="w", pady=(2, 10))
        ctk.CTkRadioButton(amb, text="Produção", variable=self.var_ambiente, value="1",
                           font=FONT_SMALL, text_color=TEXT, fg_color=ACCENT, hover_color=ACCENT2).pack(side="left", padx=(0, 14))
        ctk.CTkRadioButton(amb, text="Homologação", variable=self.var_ambiente, value="2",
                           font=FONT_SMALL, text_color=TEXT, fg_color=ACCENT, hover_color=ACCENT2).pack(side="left")

        self.btn_consultar = ctk.CTkButton(card, text="↻  Consultar NF-e", command=self._iniciar_consulta,
                                            fg_color=SUCCESS, hover_color="#145f38",
                                            text_color="white", corner_radius=14, height=44,
                                            font=("Segoe UI", 12, "bold"))
        self.btn_consultar.pack(fill="x", pady=(0, 4))

        self.btn_importar = ctk.CTkButton(
            card, text="↯  Importar por Chaves (consChNFe)", font=FONT_SMALL,
            fg_color="transparent", border_width=1, border_color=ACCENT,
            hover_color=ACCENT_LT, corner_radius=10,
            text_color=ACCENT, height=34, command=self._importar_por_extrato)
        self.btn_importar.pack(fill="x", pady=(0, 4))
        self.lbl_timer = ctk.CTkLabel(card, text="", font=FONT_SMALL, text_color=WARNING)
        self.lbl_timer.pack(anchor="w")

    def _build_panel_nsu(self, parent):
        card = self._card(parent, "CONTROLE NSU")
        self.lbl_nsu_info = ctk.CTkLabel(card, text="Último NSU: não consultado",
                                          font=FONT_SMALL, text_color=TEXT_DIM, justify="left")
        self.lbl_nsu_info.pack(anchor="w", pady=(0, 6))

        # Botão principal: zera NSU para reiniciar busca do zero
        self._btn_ghost(card, "⟳  Reiniciar busca (NSU → 0)", self._zerar_nsu)

        sep = ctk.CTkFrame(card, height=1, fg_color="#e0e0e8", corner_radius=0)
        sep.pack(fill="x", pady=8)

        self._lbl(card, "Ou definir NSU específico:")
        self.var_nsu_force = tk.StringVar(value="")
        self._entry(card, textvariable=self.var_nsu_force)
        self._btn_ghost(card, "↺  Aplicar NSU digitado", self._resetar_nsu)


    def _importar_por_extrato(self):
        """
        Importa NF-e usando consChNFe (NT 2014.002 seção 3.7).
        Usa quando distNSU retornou 0 documentos (CNPJ novo no SEFAZ).
        O usuário cola as chaves de acesso (uma por linha) — podem vir
        do extrato do portal SEFAZ ou de outro sistema.
        Limite NT: 20 consultas consChNFe por hora por CNPJ.
        """
        if not self.cert:
            messagebox.showwarning("Atenção", "Carregue o certificado antes."); return
        cnpj = self.var_cnpj.get().strip().replace(".","").replace("/","").replace("-","")
        if not cnpj or not cnpj.isdigit() or len(cnpj) != 14:
            messagebox.showwarning("Atenção", "Informe o CNPJ antes."); return

        # Janela para colar as chaves
        win = tk.Toplevel(self)
        win.title("Importar por Chaves de Acesso (consChNFe)")
        win.geometry("680x520")
        win.configure(bg=BG)
        win.grab_set()

        frame = ctk.CTkFrame(win, fg_color=BG, corner_radius=0)
        frame.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(frame, text="Chaves de Acesso (uma por linha, 44 dígitos cada):",
                     font=FONT_SMALL, text_color=TEXT).pack(anchor="w")
        ctk.CTkLabel(frame,
            text="Fonte: extrato SEFAZ (portal NF-e) · Coluna 'Chave_NF-e' do CSV/extrato",
            font=("Segoe UI", 9), text_color=TEXT_DIM).pack(anchor="w", pady=(0,6))

        txt = tk.Text(frame, height=16, font=("Courier New", 9),
                      bg=CARD, fg=TEXT, insertbackground=TEXT,
                      relief="flat", bd=0, wrap="none")
        txt.pack(fill="both", expand=True)

        def _executar():
            raw = txt.get("1.0", "end").strip()
            chaves = [c.strip() for c in raw.splitlines() if len(c.strip()) == 44 and c.strip().isdigit()]
            if not chaves:
                messagebox.showwarning("Atenção", "Nenhuma chave válida encontrada.\n"
                                       "Cada chave deve ter exatamente 44 dígitos.", parent=win)
                return
            if len(chaves) > 20:
                if not messagebox.askyesno("Atenção",
                    f"{len(chaves)} chaves detectadas.\n"
                    "A NT limita 20 consultas consChNFe/hora.\n"
                    "As primeiras 20 serão consultadas agora.\n"
                    "Continuar?", parent=win):
                    return
                chaves = chaves[:20]

            win.destroy()
            self._log(f"Importando {len(chaves)} chave(s) via consChNFe...", "info")
            self._log("(NT 2014.002 seção 3.7 — não consome token distNSU)", "info")
            self.btn_consultar.configure(state="disabled")
            self.btn_importar.configure(state="disabled")
            import threading
            threading.Thread(target=self._executar_importacao_chaves,
                             args=(cnpj, chaves), daemon=True).start()

        row = ctk.CTkFrame(frame, fg_color="transparent")
        row.pack(fill="x", pady=(8,0))
        ctk.CTkButton(row, text="Cancelar", width=100, fg_color="transparent",
                      border_width=1, border_color=TEXT_DIM, text_color=TEXT_DIM,
                      command=win.destroy).pack(side="right", padx=(4,0))
        ctk.CTkButton(row, text="Consultar Chaves", width=160,
                      fg_color=ACCENT, hover_color=ACCENT2, text_color="#fff",
                      command=_executar).pack(side="right")

    def _executar_importacao_chaves(self, cnpj: str, chaves: list):
        """Consulta cada chave via consChNFe em thread separada."""
        from src.sefaz_client import SefazClient
        import time
        try:
            ambiente  = int(self.var_ambiente.get())
            cuf_autor = self.var_uf.get().strip() or "43"
            client    = SefazClient(self.cert, ambiente, cuf_autor)
            total = erros = 0

            for i, chave in enumerate(chaves, 1):
                self._set_status(f"consChNFe {i}/{len(chaves)}...")
                self._log(f"Chave {i}/{len(chaves)}: {chave[:20]}...", "info")

                r = client.consultar_por_chave(cnpj, chave)

                if r["status"] == "consumo_indevido":
                    self._log("⚠ Limite de 20 consultas/hora atingido (cStat=656).", "aviso")
                    self._log(f"  {i-1} chaves importadas. Tente as restantes em 1 hora.", "aviso")
                    break
                elif r["status"] == "erro":
                    self._log(f"  ✘ {r['mensagem']}", "erro")
                    erros += 1
                elif r["status"] in ("ok", "sem_novos"):
                    nfes = r.get("notas", [])
                    if nfes:
                        self.storage.salvar_notas(cnpj, nfes)
                        total += len(nfes)
                        ult_nsu = r.get("ult_nsu", 0)
                        self.after(0, self._atualizar_resultados, nfes, ult_nsu, cnpj)
                        sit = nfes[0].get("situacao","?")
                        schema = nfes[0].get("schema","?")
                        tipo = "XML completo (procNFe)" if "procNFe" in schema else "Resumo (resNFe — faça Manifestação)"
                        self._log(f"  ✔ {sit} · {tipo}", "ok")
                    else:
                        self._log(f"  {r['mensagem']}", "aviso")

                # NT 3.11.4.2: delay entre consultas consChNFe
                if i < len(chaves):
                    time.sleep(2.0)

            self._log(f"─── Importação concluída: {total} nota(s) · {erros} erro(s) ───",
                      "ok" if total > 0 else "aviso")
            if total > 0:
                self._log("► Para obter XML completo: faça Manifestação (Ciência) em cada nota.", "info")
            self._set_status(f"Importação: {total} nota(s)")

        except Exception as e:
            import traceback
            self._log(f"Erro: {e}", "erro")
            self._log(traceback.format_exc()[:400], "erro")
        finally:
            self.after(0, self._finalizar_consulta)
    def _build_panel_resultados(self, parent):
        title_row = ctk.CTkFrame(parent, fg_color="transparent", corner_radius=0)
        title_row.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(title_row, text="NOTAS FISCAIS ENCONTRADAS", font=FONT_HEAD, text_color=TEXT).pack(side="left")
        self.lbl_total_nfe = ctk.CTkLabel(title_row, text="", font=FONT_SMALL, text_color=TEXT_DIM)
        self.lbl_total_nfe.pack(side="right")

        tv_frame = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=16,
                                 border_width=2, border_color="#eef0f8")
        tv_frame.pack(fill="both", expand=True, pady=(0, 8))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Clay.Treeview", background=CARD, foreground=TEXT,
                         fieldbackground=CARD, borderwidth=0, font=FONT_MONO, rowheight=26)
        style.configure("Clay.Treeview.Heading", background=DARK, foreground="white",
                         font=("Segoe UI", 9, "bold"), borderwidth=0, relief="flat")
        style.map("Clay.Treeview", background=[("selected", ACCENT_LT)],
                  foreground=[("selected", ACCENT2)])

        cols = ("sel", "nsu", "chave", "emitente", "valor", "emissao", "situacao", "xml")
        self.tree = ttk.Treeview(tv_frame, columns=cols, show="headings", style="Clay.Treeview")
        headers = {
            "sel": ("✓", 30), "nsu": ("NSU", 80), "chave": ("Chave de Acesso", 330),
            "emitente": ("Emitente / CNPJ", 180), "valor": ("Valor (R$)", 100),
            "emissao": ("Emissão", 120), "situacao": ("Situação", 90), "xml": ("XML", 45),
        }
        for col, (lbl, w) in headers.items():
            anc = "center" if col in ("sel", "xml") else "w"
            self.tree.heading(col, text=lbl,
                              command=lambda c=col: self._toggle_sel_all() if c == "sel" else None)
            self.tree.column(col, width=w, anchor=anc, minwidth=w)

        sy = ttk.Scrollbar(tv_frame, orient="vertical",   command=self.tree.yview)
        sx = ttk.Scrollbar(tv_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        sy.pack(side="right",  fill="y")
        sx.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True, padx=4, pady=4)
        self.tree.bind("<Double-1>", self._detalhar_nfe)
        self.tree.bind("<Button-1>", self._toggle_sel_item)

        # Barra de ações
        bar = ctk.CTkFrame(parent, fg_color=INPUT_BG, corner_radius=14,
                            border_width=2, border_color=BORDER)
        bar.pack(fill="x", pady=(0, 6))

        brow1 = ctk.CTkFrame(bar, fg_color="transparent", corner_radius=0)
        brow1.pack(fill="x", padx=10, pady=(8, 4))

        def ab(txt, cmd, fg=None, tc="white"):
            ctk.CTkButton(brow1, text=txt, command=cmd,
                          fg_color=fg or INPUT_BG, hover_color=ACCENT_LT if not fg else ACCENT2,
                          text_color=tc if fg else TEXT, corner_radius=10, height=32,
                          font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 5))

        ab("Detalhes",              lambda: self._detalhar_nfe(None))
        ab("⬇ Baixar XML",          self._baixar_xml,       fg=ACCENT)
        ab("⬇ Baixar Selecionados", self._baixar_xml_lote,  fg=ACCENT2)
        ab("Exportar JSON",         self._exportar_json)
        ctk.CTkButton(brow1, text="Limpar", command=self._limpar_lista,
                      fg_color=ERROR_LT, hover_color="#ffd0d0", text_color=ERROR,
                      corner_radius=10, height=32, font=("Segoe UI", 9, "bold")).pack(side="right")

        brow2 = ctk.CTkFrame(bar, fg_color="transparent", corner_radius=0)
        brow2.pack(fill="x", padx=10, pady=(0, 8))
        ctk.CTkLabel(brow2, text="Manifestar:", font=FONT_BADGE, text_color=TEXT_DIM).pack(side="left", padx=(0, 8))

        for lbl, tp, fg, tc in [
            ("Ciência",         "210210", "#e8f8ff",  "#0070a0"),
            ("Confirmação",     "210200", SUCCESS_LT, SUCCESS),
            ("Desconhecimento", "210220", "#fff8e0",  WARNING),
            ("Não Realizada",   "210240", ERROR_LT,   ERROR),
        ]:
            tp_l = tp
            ctk.CTkButton(brow2, text=lbl, command=lambda t=tp_l: self._manifestar(t),
                          fg_color=fg, hover_color=ACCENT_LT, text_color=tc,
                          corner_radius=20, height=28,
                          font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 5))

    def _build_panel_log(self, parent):
        log_card = ctk.CTkFrame(parent, fg_color=LOG_BG, corner_radius=16)
        log_card.pack(fill="x")

        hrow = ctk.CTkFrame(log_card, fg_color="transparent", corner_radius=0)
        hrow.pack(fill="x", padx=14, pady=(10, 4))
        pill = ctk.CTkFrame(hrow, fg_color="#252540", corner_radius=20)
        pill.pack(side="left")
        ctk.CTkLabel(pill, text="  LOG DE EVENTOS  ", font=FONT_BADGE, text_color="#7dd3fc").pack(padx=4, pady=4)
        ctk.CTkButton(hrow, text="Limpar", command=self._limpar_log, width=60, height=26,
                      font=FONT_SMALL, fg_color="#252540", hover_color="#303050",
                      text_color="#8899aa", corner_radius=8).pack(side="right")

        self.log_text = scrolledtext.ScrolledText(log_card, height=7, font=FONT_MONO,
                                                   bg=LOG_BG, fg="#c8d8f0",
                                                   insertbackground=ACCENT,
                                                   relief="flat", bd=0, state="disabled")
        self.log_text.pack(fill="x", padx=10, pady=(0, 10))
        self.log_text.tag_configure("ok",    foreground="#4ade80")
        self.log_text.tag_configure("erro",  foreground="#f87171")
        self.log_text.tag_configure("aviso", foreground="#fbbf24")
        self.log_text.tag_configure("info",  foreground="#7dd3fc")
        self.log_text.tag_configure("ts",    foreground="#475569")

    def _build_status_bar(self):
        bar = ctk.CTkFrame(self, fg_color=DARK, corner_radius=0, height=36)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        self.lbl_status = ctk.CTkLabel(bar, text="Aguardando certificado...",
                                        font=("Segoe UI", 9), text_color="#8899cc")
        self.lbl_status.pack(side="left", padx=16)
        vf = ctk.CTkFrame(bar, fg_color=ACCENT, corner_radius=10)
        vf.pack(side="right", padx=16, pady=6)
        ctk.CTkLabel(vf, text="  Abbas · v1.0  ", font=FONT_BADGE, text_color="white").pack(padx=4, pady=2)
        self.progress = ttk.Progressbar(bar, mode="indeterminate", length=100)
        self.progress.pack(side="right", padx=8)

    # ── AÇÕES ─────────────────────────────────────────────────────────────────

    def _selecionar_cert(self):
        path = filedialog.askopenfilename(
            title="Selecionar Certificado Digital",
            filetypes=[("Certificado Digital", "*.pfx *.p12"), ("Todos", "*.*")])
        if path:
            self.var_cert_path.set(os.path.basename(path))
            self._cert_path_full = path

    def _carregar_cert(self):
        path  = getattr(self, "_cert_path_full", None)
        senha = self.var_senha.get()
        if not path:
            self._log("Selecione um arquivo de certificado primeiro.", "aviso"); return
        if not senha:
            self._log("Informe a senha do certificado.", "aviso"); return
        try:
            self.cert = CertificadoDigital(path, senha)
            info = self.cert.info()
            self._atualizar_status_cert(ok=True, info=info)
            self._log(f"Certificado: {info['titular']} | CNPJ: {info['cnpj']} | Val: {info['validade']}", "ok")
            if info["cnpj"]: self.var_cnpj.set(info["cnpj"])
            self.storage.salvar_cert_config("pfx", path, info["cnpj"])
        except Exception as e:
            self.cert = None
            self._atualizar_status_cert(ok=False)
            self._log(f"Erro: {e}", "erro")

    def _restaurar_cert_salvo(self):
        cfg = self.storage.get_cert_config()
        if not cfg: return
        tipo, valor, cnpj = cfg.get("tipo",""), cfg.get("valor",""), cfg.get("cnpj","")
        if tipo == "pfx" and valor and os.path.exists(valor):
            self.var_cert_path.set(os.path.basename(valor))
            self._cert_path_full = valor
            if cnpj: self.var_cnpj.set(cnpj)
            self._log(f"Certificado anterior: {os.path.basename(valor)} — informe a senha e clique em Carregar.", "info")
        elif tipo == "windows" and valor:
            for c in listar_certificados_windows():
                if c["titular"] == valor:
                    self._carregar_cert_do_windows(c); return

    def _selecionar_cert_windows(self):
        self._log("Lendo certificados do Windows...", "info")
        certs = listar_certificados_windows()
        if not certs:
            messagebox.showwarning("Nenhum certificado",
                "Não foram encontrados certificados válidos no store Pessoal."); return

        win = ctk.CTkToplevel(self)
        win.title("Certificados instalados no Windows")
        win.geometry("640x360")
        win.configure(fg_color=BG)
        win.grab_set()
        ctk.CTkLabel(win, text="Selecione o certificado:", font=FONT_HEAD, text_color=TEXT).pack(padx=20, pady=(16,8), anchor="w")

        lb_f = ctk.CTkFrame(win, fg_color=CARD, corner_radius=12)
        lb_f.pack(fill="both", expand=True, padx=20, pady=(0,10))
        lb = tk.Listbox(lb_f, font=FONT_BODY, bg=CARD, fg=TEXT,
                        selectbackground=ACCENT, selectforeground="white",
                        relief="flat", bd=0, activestyle="none")
        lb.pack(fill="both", expand=True, padx=4, pady=4)
        for c in certs:
            lb.insert("end", f"  {c['titular']}  |  {c['cnpj'] or 'CPF/CNPJ não extraído'}  |  Val: {c['validade']}")
        if certs: lb.selection_set(0)

        def confirmar():
            sel = lb.curselection()
            if not sel: return
            win.destroy()
            self._carregar_cert_do_windows(certs[sel[0]])

        br = ctk.CTkFrame(win, fg_color="transparent", corner_radius=0)
        br.pack(fill="x", padx=20, pady=(0,16))
        ctk.CTkButton(br, text="Usar este certificado", command=confirmar,
                      fg_color=ACCENT, hover_color=ACCENT2, corner_radius=12, height=38, font=FONT_HEAD).pack(side="left")
        ctk.CTkButton(br, text="Cancelar", command=win.destroy,
                      fg_color=INPUT_BG, hover_color=ACCENT_LT, text_color=TEXT_DIM,
                      corner_radius=12, height=38, font=FONT_SMALL).pack(side="left", padx=(8,0))
        lb.bind("<Double-1>", lambda e: confirmar())

    def _carregar_cert_do_windows(self, cert_info):
        self._log(f"Carregando: {cert_info['titular']}...", "info")
        try:
            from src.win_cert_store import exportar_pem_do_store
            cert_pem, key_pem = exportar_pem_do_store(cert_info["der"])
            self.cert = _CertificadoWindowsStore(cert_pem, key_pem, cert_info)
        except Exception as e_exp:
            self._log(f"Exportação bloqueada — modo WinHTTP...", "aviso")
            try:
                import win32com.client
                self.cert = _CertificadoWinHTTP(cert_info)
            except ImportError:
                self._log("pywin32 não instalado: pip install pywin32", "erro"); return
            except Exception as e:
                self._log(f"Erro WinHTTP: {e}", "erro"); return
        info = self.cert.info()
        self._atualizar_status_cert(ok=True, info=info)
        self._log(f"Carregado: {info['titular']} | CNPJ: {info['cnpj']}", "ok")
        if info["cnpj"]: self.var_cnpj.set(info["cnpj"])
        self.storage.salvar_cert_config("windows", cert_info["titular"], info["cnpj"])

    def _atualizar_status_cert(self, ok=None, info=None):
        if ok is True and info:
            self.lbl_cert_status.configure(
                text=f"✔ {info['titular'][:35]}  |  Val: {info['validade']}", text_color=SUCCESS)
            self._set_status(f"Certificado · {info['titular'][:40]}")
        elif ok is False:
            self.lbl_cert_status.configure(text="✘ Falha ao carregar", text_color=ERROR)
        else:
            self.lbl_cert_status.configure(text="Nenhum certificado carregado", text_color=TEXT_DIM)

    def _iniciar_consulta(self):
        if not self.cert:
            messagebox.showwarning("Atenção", "Carregue o certificado primeiro."); return
        cnpj = self.var_cnpj.get().strip().replace(".","").replace("/","").replace("-","")
        if len(cnpj) != 14 or not cnpj.isdigit():
            messagebox.showwarning("Atenção", "CNPJ inválido."); return
        self.btn_consultar.configure(state="disabled")
        self.progress.start(10)
        self._log(f"Iniciando consulta para CNPJ {cnpj}...", "info")
        threading.Thread(target=self._executar_consulta, args=(cnpj,), daemon=True).start()

    def _executar_consulta(self, cnpj):
        """
        Loop de consulta completo conforme NT 2014.002:
        - Timeout 90s por chamada (SEFAZ pode demorar 30-45s)
        - Delay 1s entre lotes (evita cStat 656)
        - Loop enquanto cStat=138 (tem_mais=True)
        - Log com descrição humana de cada cStat
        - Disponibilidade: 90 dias a partir da autorização
        """
        from src.sefaz_client import DELAY_ENTRE_LOTES
        import time
        try:
            ambiente  = int(self.var_ambiente.get())
            cuf_autor = self.var_uf.get().strip() or "43"
            client    = SefazClient(self.cert, ambiente, cuf_autor)
            total      = 0
            rodada     = 0
            ultimo_nsu = self.storage.get_nsu(cnpj)

            self._log(f"─── Iniciando busca completa (até 90 dias) ───", "info")
            self._log(f"Ambiente: {'Produção' if ambiente==1 else 'Homologação'} · UF: {cuf_autor}", "info")

            # NT 2014.002 seção 3.4: NSU=0 é válido.
            # Novo usuário: 1ª consulta retorna cStat=137 (NORMAL).
            # O SEFAZ começa a gerar NSU a partir daí.
            # NÃO fazer consultas extras — cada uma consome o token de 60min.
            if ultimo_nsu == 0:
                self._log("NSU=0 — primeira consulta para este CNPJ.", "info")
                self._log("(NT 2014.002 seção 3.4: novo usuário recebe cStat=137 na 1ª consulta)", "info")

            self._log(f"NSU de partida: {ultimo_nsu:015d}", "info")
            self._log(f"Timeout por chamada: 90s · Delay entre lotes: {DELAY_ENTRE_LOTES}s", "info")
            self._log(f"─────────────────────────────────────────────", "info")

            while True:
                rodada += 1
                self._set_status(f"Lote {rodada} — NSU {ultimo_nsu:015d} — {total} docs até agora...")
                self.after(0, lambda r=rodada, n=ultimo_nsu, t=total:
                    self._log(f"Lote {r} · Consultando NSU a partir de {n:015d} ({t} docs acumulados)...", "info"))

                r = client.consultar_distribuicao(cnpj, ultimo_nsu)

                # ── Trata cada situação com mensagem clara ──────────────────
                if r["status"] == "consumo_indevido":
                    self._log(f"⚠ {r['mensagem']}", "aviso")
                    self._log("O SEFAZ limita consultas a 1 por hora por CNPJ.", "aviso")
                    self._log("Aguarde 60 minutos e consulte novamente.", "aviso")
                    self.after(0, self._iniciar_countdown)
                    return

                if r["status"] == "erro":
                    self._log(f"✘ Erro na comunicação com SEFAZ:", "erro")
                    for linha in r["mensagem"].split("\n"):
                        if linha.strip():
                            self._log(f"  {linha}", "erro")
                    codigo = r.get("codigo","?")
                    if codigo == "341":
                        self._log("  → Verifique o campo 'Cód. UF (IBGE)' — deve ser o código da sua UF (43=RS, 35=SP...)", "aviso")
                    elif codigo in ("214","302"):
                        self._log("  → O CNPJ do certificado não bate com o CNPJ informado.", "aviso")
                    elif codigo in ("WINHTTP_VAZIO","CSTAT_VAZIO","SEM_RETORNO"):
                        self._log("  → Verifique o arquivo logs/diag_*.txt para diagnóstico completo.", "aviso")
                        self._log("  → Tente usar o arquivo .pfx diretamente ao invés do certificado do Windows.", "aviso")
                    elif codigo == "656":
                        self._log("  → Aguarde 60 minutos.", "aviso")
                        self.after(0, self._iniciar_countdown)
                    return

                # ── Processa retorno conforme NT 2014.002 seções 3.5 e 3.11.4 ──
                # ultNSU = até onde o SEFAZ processou — usar na PRÓXIMA chamada
                # maxNSU = teto do AN para este CNPJ — NÃO usar como próximo NSU
                nfes    = r.get("notas", [])
                ult_nsu = r.get("ult_nsu", ultimo_nsu)
                max_nsu = r.get("max_nsu", ultimo_nsu)
                tem_mais = r.get("tem_mais", False)

                # Persiste ultNSU (ponto de continuação para próxima sessão)
                if ult_nsu > ultimo_nsu:
                    self.storage.set_nsu(cnpj, ult_nsu)

                if nfes:
                    self.storage.salvar_notas(cnpj, nfes)
                    total += len(nfes)
                    self.after(0, self._atualizar_resultados, nfes, ult_nsu, cnpj)
                    tipos = {}
                    for n in nfes:
                        sit = n.get("situacao","?")
                        tipos[sit] = tipos.get(sit,0)+1
                    resumo = " | ".join(f"{v}x {k}" for k,v in tipos.items())
                    self._log(f"  ✔ {len(nfes)} doc(s) recebidos — {resumo}", "ok")
                else:
                    if tem_mais:
                        self._log(f"  Sem docs neste bloco · Avançando NSU para {ult_nsu:015d}...", "info")
                    else:
                        self._log(f"  {r['mensagem']}", "info")

                self._log(
                    f"  ultNSU: {ult_nsu:015d} · maxNSU: {max_nsu:015d} · "
                    f"Acumulado: {total} · {'Continuando...' if tem_mais else 'Fim'}", "info")

                # ── Decide se continua o loop ───────────────────────────────
                if not tem_mais:
                    if total == 0 and r["status"] == "sem_novos":
                        if ultimo_nsu == 0:
                            # NT 3.4 v1.10: novo usuário — comportamento esperado
                            self._log("", "info")
                            self._log("► CNPJ novo no SEFAZ — comportamento NORMAL (NT 2014.002 seção 3.4):", "aviso")
                            self._log("  O SEFAZ passou a gerar NSU para este CNPJ a partir desta consulta.", "aviso")
                            self._log("  Aguarde 60 minutos e consulte novamente.", "aviso")
                            self._log("  Na próxima consulta os documentos dos últimos 90 dias aparecerão.", "aviso")
                        else:
                            self._log("Nenhum documento novo encontrado.", "info")
                            self._log("O SEFAZ está sincronizado — não há NF-e desde a última consulta.", "info")
                    break

                if ult_nsu <= ultimo_nsu:
                    self._log("  ⚠ ultNSU não avançou — interrompendo para evitar loop infinito.", "aviso")
                    break

                ultimo_nsu = ult_nsu
                time.sleep(DELAY_ENTRE_LOTES)

            # ── Sumário final ───────────────────────────────────────────────
            self._log(f"─── Consulta finalizada ───", "info")
            self._log(f"Total: {total} documento(s) em {rodada} lote(s)", "ok" if total>0 else "aviso")
            self._log(f"NSU atual: {ultimo_nsu:015d}", "info")
            if total > 0:
                self._log(f"► Faça a Manifestação (Ciência) para liberar o XML completo de cada nota.", "info")
            self._set_status(f"Concluído · {total} doc(s)")
            self.after(0, self._atualizar_nsu_label)

        except Exception as e:
            import traceback
            self._log(f"Erro inesperado: {e}", "erro")
            self._log(traceback.format_exc()[:500], "erro")
        finally:
            self.after(0, self._finalizar_consulta)

    def _finalizar_consulta(self):
        self.progress.stop()
        self.btn_consultar.configure(state="normal")
        self._atualizar_nsu_label()

    def _atualizar_resultados(self, nfes, novo_nsu, cnpj):
        for nfe in nfes: self._inserir_nfe_tree(nfe)
        self.lbl_total_nfe.configure(text=f"{len(self.tree.get_children())} nota(s)")
        self._atualizar_nsu_label()

    def _inserir_nfe_tree(self, nfe):
        nsu_str = str(nfe.get("nsu","")).strip()
        chave   = nfe.get("chave","").strip()
        tipo    = nfe.get("tipo","")
        for item in self.tree.get_children():
            if self.tree.item(item,"values")[1].strip() == nsu_str: return
        sit = nfe.get("situacao","")
        cor = ("ok" if sit=="Autorizada" else
               "cancel" if sit in ("Cancelada","Denegada","Cancelamento") else
               "evento" if "Evento" in tipo else "")
        emitente = nfe.get("emitente","") or (tipo if "Evento" in tipo else "")
        self.tree.insert("","end", values=(
            "", nsu_str, chave, emitente,
            nfe.get("valor",""), nfe.get("emissao",""),
            sit, "✔" if nfe.get("xml_raw","") else ""), tags=(cor,))
        self.tree.tag_configure("ok",     foreground=SUCCESS)
        self.tree.tag_configure("cancel", foreground=ERROR)
        self.tree.tag_configure("evento", foreground=WARNING)

    def _toggle_sel_all(self):
        children = self.tree.get_children()
        novo = "☑" if any(self.tree.item(i,"values")[0]=="" for i in children) else ""
        for item in children:
            v = list(self.tree.item(item,"values")); v[0]=novo; self.tree.item(item,values=v)

    def _toggle_sel_item(self, event):
        if self.tree.identify_region(event.x,event.y)=="cell" and self.tree.identify_column(event.x)=="#1":
            item = self.tree.identify_row(event.y)
            if item:
                v = list(self.tree.item(item,"values"))
                v[0] = "" if v[0]=="☑" else "☑"
                self.tree.item(item,values=v)
            return "break"

    def _get_selecionados(self):
        cnpj = self.var_cnpj.get().strip().replace(".","").replace("/","").replace("-","")
        return [self.storage.get_nota(cnpj, v[2].strip()) or {"nsu":int(v[1].strip()),"chave":v[2].strip()}
                for v in (self.tree.item(i,"values") for i in self.tree.get_children()) if v[0]=="☑"]

    def _atualizar_nsu_label(self):
        cnpj = self.var_cnpj.get().strip().replace(".","").replace("/","").replace("-","")
        if cnpj and cnpj.isdigit() and len(cnpj)==14:
            self.lbl_nsu_info.configure(text=f"CNPJ: {cnpj}\nÚltimo NSU: {self.storage.get_nsu(cnpj):015d}")

    def _zerar_nsu(self):
        """Zera o NSU para 0 — reinicia a busca do início (NT 2014.002 seção 3.4)."""
        cnpj = self.var_cnpj.get().strip().replace(".","").replace("/","").replace("-","")
        if not cnpj or not cnpj.isdigit() or len(cnpj) != 14:
            messagebox.showwarning("Atenção", "Informe o CNPJ antes de zerar o NSU."); return
        if messagebox.askyesno("Confirmar",
            f"Zerar o NSU para 0 do CNPJ {cnpj}?\n\n"
            "A próxima consulta buscará documentos dos últimos 90 dias.\n"
            "Se for a primeira consulta deste CNPJ, aguarde 60 min após zerar."):
            self.storage.set_nsu(cnpj, 0)
            self._atualizar_nsu_label()
            self._log("NSU zerado para 0 — próxima consulta buscará do início.", "aviso")
            self._log("ATENÇÃO: Se cStat=137 na próxima consulta → aguarde 60 minutos.", "aviso")

    def _resetar_nsu(self):
        cnpj = self.var_cnpj.get().strip().replace(".","").replace("/","").replace("-","")
        if not cnpj or not cnpj.isdigit() or len(cnpj) != 14:
            messagebox.showwarning("Atenção","CNPJ inválido."); return
        val = self.var_nsu_force.get().strip()
        if not val:
            messagebox.showwarning("Atenção","Digite um NSU no campo acima."); return
        try:
            nsu = int(val)
            self.storage.set_nsu(cnpj, nsu)
            self._atualizar_nsu_label()
            self._log(f"NSU definido para {nsu:015d}", "aviso")
        except ValueError:
            messagebox.showwarning("Atenção","NSU inválido — use apenas números.")

    def _iniciar_countdown(self):
        s = [3600]
        def tick():
            if s[0]>0:
                self.lbl_timer.configure(text=f"⏱ {s[0]//60:02d}:{s[0]%60:02d}")
                s[0]-=1; self.after(1000,tick)
            else:
                self.lbl_timer.configure(text="✔ Pronto")
                self.btn_consultar.configure(state="normal")
        self.btn_consultar.configure(state="disabled"); tick()

    def _detalhar_nfe(self, event):
        sel = self.tree.selection()
        if not sel: return
        vals  = self.tree.item(sel[0],"values")
        cnpj  = self.var_cnpj.get().strip().replace(".","").replace("/","").replace("-","")
        chave = vals[2].strip()
        nota  = self.storage.get_nota(cnpj, chave)
        win   = ctk.CTkToplevel(self)
        win.title(f"Detalhes · {chave[:12]}...")
        win.geometry("720x540")
        win.configure(fg_color=BG)
        win.grab_set()
        ctk.CTkLabel(win, text="Detalhes da NF-e", font=FONT_TITLE, text_color=TEXT).pack(padx=20,pady=(16,8),anchor="w")
        txt = scrolledtext.ScrolledText(win, font=FONT_MONO, bg=LOG_BG, fg="#c8d8f0", relief="flat", bd=0)
        txt.pack(fill="both",expand=True,padx=20,pady=(0,16))
        txt.insert("end", json.dumps(nota,indent=2,ensure_ascii=False) if nota else
                   f"NSU: {vals[1]}\nChave: {chave}\nEmitente: {vals[3]}\nValor: {vals[4]}\nEmissão: {vals[5]}\nSituação: {vals[6]}")
        txt.configure(state="disabled")

    def _baixar_xml(self):
        sel = self.tree.selection()
        if not sel: messagebox.showwarning("Atenção","Selecione uma NF-e."); return
        if not self.cert: messagebox.showwarning("Atenção","Carregue o certificado."); return
        vals = self.tree.item(sel[0],"values")
        nsu_str, chave = vals[1].strip(), vals[2].strip()
        try: nsu = int(nsu_str)
        except ValueError: messagebox.showerror("Erro",f"NSU inválido: {nsu_str}"); return
        cnpj = self.var_cnpj.get().strip().replace(".","").replace("/","").replace("-","")
        nota = self.storage.get_nota(cnpj, chave)
        if nota and nota.get("xml_raw"): self._salvar_xml_dialog(nota["xml_raw"], chave); return
        self._log(f"Baixando XML · NSU {nsu:015d}...", "info")
        self.progress.start(10)
        def worker():
            try:
                r = EventoClient(self.cert, int(self.var_ambiente.get())).baixar_xml(cnpj, nsu)
                if r["status"]=="ok" and r.get("xml_str"):
                    xml = r["xml_str"]
                    if nota: nota["xml_raw"]=xml; self.storage.salvar_notas(cnpj,[nota])
                    self._log(f"XML baixado · {len(xml)} bytes","ok")
                    self.after(0, self._salvar_xml_dialog, xml, chave)
                    self.after(0, self._marcar_xml_baixado, nsu)
                elif r["status"]=="sem_xml":
                    self._log("XML indisponível — faça Ciência primeiro.","aviso")
                    self.after(0, messagebox.showwarning, "Indisponível", "Faça Manifestação → Ciência primeiro.")
                else: self._log(f"Erro: {r.get('mensagem','')}","erro")
            except Exception as e: self._log(f"Erro: {e}","erro")
            finally: self.after(0, self.progress.stop)
        threading.Thread(target=worker,daemon=True).start()

    def _salvar_xml_dialog(self, xml_str, chave):
        path = filedialog.asksaveasfilename(defaultextension=".xml",
            filetypes=[("XML","*.xml")], initialfile=f"NFe_{chave[:20]}.xml")
        if path:
            open(path,"w",encoding="utf-8").write(xml_str)
            self._log(f"Salvo: {path}","ok")

    def _manifestar(self, tp_evento):
        sel = self.tree.selection()
        if not sel: messagebox.showwarning("Atenção","Selecione uma NF-e."); return
        if not self.cert: messagebox.showwarning("Atenção","Carregue o certificado."); return
        vals  = self.tree.item(sel[0],"values")
        chave = vals[2].strip()
        cnpj  = self.var_cnpj.get().strip().replace(".","").replace("/","").replace("-","")
        desc  = TIPOS_EVENTO.get(tp_evento, tp_evento)
        justificativa = ""
        if tp_evento == "210240":
            win = ctk.CTkToplevel(self); win.title("Justificativa")
            win.geometry("440x160"); win.configure(fg_color=BG); win.grab_set()
            ctk.CTkLabel(win,text="Justificativa (15–255 caracteres):",font=FONT_SMALL,text_color=TEXT_DIM).pack(padx=16,pady=(14,4),anchor="w")
            vj = tk.StringVar(); ctk.CTkEntry(win,textvariable=vj,fg_color=INPUT_BG,border_color=BORDER,corner_radius=10).pack(fill="x",padx=16,pady=(0,10))
            res=[None]
            def ok():
                j=vj.get().strip()
                if len(j)<15: messagebox.showwarning("Atenção","Mínimo 15 caracteres.",parent=win); return
                res[0]=j; win.destroy()
            ctk.CTkButton(win,text="Confirmar",command=ok,fg_color=ACCENT,hover_color=ACCENT2,corner_radius=12,font=FONT_HEAD).pack(padx=16,fill="x")
            win.wait_window()
            if res[0] is None: return
            justificativa = res[0]
        if not messagebox.askyesno("Confirmar",f"Evento: {desc}\nChave: {chave}\n\nConfirmar?"): return
        self._log(f"Enviando [{tp_evento}] {desc} · {chave[:20]}...", "info")
        self.progress.start(10)
        def worker():
            try:
                r = EventoClient(self.cert, int(self.var_ambiente.get())).manifestar(cnpj, chave, tp_evento, justificativa)
                if r["status"]=="ok":
                    self._log(f"Registrado! Protocolo: {r.get('protocolo','')}","ok")
                    self.after(0, self._atualizar_situacao_tree, chave, desc)
                    nota = self.storage.get_nota(cnpj, chave)
                    if nota:
                        nota["situacao"]=desc; nota["protocolo_manifestacao"]=r.get("protocolo","")
                        self.storage.salvar_notas(cnpj,[nota])
                else: self._log(f"Erro: [{r['codigo']}] {r['mensagem']}","erro")
            except Exception as e: self._log(f"Erro: {e}","erro")
            finally: self.after(0, self.progress.stop)
        threading.Thread(target=worker,daemon=True).start()

    def _atualizar_situacao_tree(self, chave, nova_sit):
        for item in self.tree.get_children():
            v = list(self.tree.item(item,"values"))
            if v[2].strip()==chave: v[6]=nova_sit; self.tree.item(item,values=v); break

    def _baixar_xml_lote(self):
        sels = self._get_selecionados()
        if not sels: messagebox.showwarning("Atenção","Marque notas com ✓."); return
        if not self.cert: messagebox.showwarning("Atenção","Carregue o certificado."); return
        pasta = filedialog.askdirectory(title="Pasta para salvar XMLs")
        if not pasta: return
        cnpj = self.var_cnpj.get().strip().replace(".","").replace("/","").replace("-","")
        self._log(f"Lote: {len(sels)} nota(s)...", "info")
        self.progress.start(10)
        self.btn_consultar.configure(state="disabled")
        def worker():
            ok=erros=0; client=EventoClient(self.cert,int(self.var_ambiente.get()))
            for nota in sels:
                nsu=nota.get("nsu",0); chave=nota.get("chave","").strip()
                xml=nota.get("xml_raw","")
                if not xml and nsu:
                    try:
                        r=client.baixar_xml(cnpj,int(nsu))
                        if r.get("status")=="ok": xml=r["xml_str"]; nota["xml_raw"]=xml; self.storage.salvar_notas(cnpj,[nota])
                    except Exception as e: self._log(f"Erro NSU {nsu}: {e}","erro"); erros+=1; continue
                if xml:
                    open(os.path.join(pasta,f"NFe_{chave or nsu}.xml"),"w",encoding="utf-8").write(xml)
                    self._log(f"✔ NFe_{chave or nsu}.xml","ok"); self.after(0,self._marcar_xml_baixado,nsu); ok+=1
                else: self._log(f"✘ NSU {nsu} — sem XML (faça Ciência)","aviso"); erros+=1
            self._log(f"Lote: {ok} salvo(s), {erros} sem XML.","ok" if not erros else "aviso")
            self.after(0,self.progress.stop); self.after(0,lambda:self.btn_consultar.configure(state="normal"))
        threading.Thread(target=worker,daemon=True).start()

    def _marcar_xml_baixado(self, nsu):
        nsu_str=str(nsu).strip()
        for item in self.tree.get_children():
            v=list(self.tree.item(item,"values"))
            if v[1].strip()==nsu_str: v[7]="✔"; self.tree.item(item,values=v); break

    def _exportar_json(self):
        cnpj = self.var_cnpj.get().strip().replace(".","").replace("/","").replace("-","")
        notas = self.storage.get_todas_notas(cnpj)
        if not notas: messagebox.showinfo("Exportar","Nenhuma nota."); return
        path = filedialog.asksaveasfilename(defaultextension=".json",filetypes=[("JSON","*.json")],
            initialfile=f"nfe_{cnpj}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
        if path:
            json.dump(notas,open(path,"w",encoding="utf-8"),indent=2,ensure_ascii=False)
            self._log(f"Exportado: {path}","ok")

    def _carregar_nfes_salvas(self):
        cnpj = self.var_cnpj.get().strip()
        for nota in (self.storage.get_todas_notas(cnpj) if cnpj else []):
            self._inserir_nfe_tree(nota)

    def _limpar_lista(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        self.lbl_total_nfe.configure(text="")

    def _limpar_log(self):
        self.log_text.configure(state="normal"); self.log_text.delete("1.0","end"); self.log_text.configure(state="disabled")

    def _log(self, msg, tipo="info"):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] ", "ts")
        self.log_text.insert("end", f"{msg}\n", tipo)
        self.log_text.see("end"); self.log_text.configure(state="disabled")

    def _set_status(self, msg):
        self.after(0, lambda: self.lbl_status.configure(text=msg))


# ── Classes auxiliares ────────────────────────────────────────────────────────

class _CertificadoWinHTTP:
    def __init__(self, ci):
        self._titular=ci.get("titular",""); self._cnpj=ci.get("cnpj","")
        self._validade=ci.get("validade",""); self._winhttp_cn=self._titular
    def info(self): return {"titular":self._titular,"cnpj":self._cnpj,"validade":self._validade,"arquivo":"Windows Store (WinHTTP)"}
    def exportar_pem_temp(self): raise RuntimeError("Não-exportável — use WinHTTP.")
    @property
    def titular(self): return self._titular
    @property
    def cnpj(self): return self._cnpj
    @property
    def validade(self): return self._validade


class _CertificadoWindowsStore:
    def __init__(self, cert_pem, key_pem, ci):
        self._cert_pem=cert_pem; self._key_pem=key_pem
        self._titular=ci.get("titular",""); self._cnpj=ci.get("cnpj",""); self._validade=ci.get("validade","")
    def info(self): return {"titular":self._titular,"cnpj":self._cnpj,"validade":self._validade,"arquivo":"Windows Store"}
    def exportar_pem_temp(self):
        import tempfile
        cf=tempfile.NamedTemporaryFile(delete=False,suffix=".cert.pem"); cf.write(self._cert_pem); cf.close()
        kf=tempfile.NamedTemporaryFile(delete=False,suffix=".key.pem");  kf.write(self._key_pem);  kf.close()
        return cf.name, kf.name
    @property
    def titular(self): return self._titular
    @property
    def cnpj(self): return self._cnpj
    @property
    def validade(self): return self._validade
