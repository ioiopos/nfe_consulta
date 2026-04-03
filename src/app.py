"""
Consulta NF-e Destinadas - Interface Principal
Consulta NF-e Modelo 55 via Web Service NFeDistribuicaoDFe (SEFAZ)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import json
import os
import sys
from datetime import datetime
from pathlib import Path

# Adiciona o diretório pai ao path para imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.certificado import CertificadoDigital
from src.sefaz_client import SefazClient
from src.storage import StorageNSU
from src.parser import ParserNFe
from src.evento_client import EventoClient, TIPOS_EVENTO
from src.win_cert_store import listar_certificados_windows, IS_WINDOWS

# ─────────────────────────────────────────────────────────────────────────────
# CORES — Abbas Driver Portal, fiel ao portal real
# Navbar/header: --dark #1a1a2e  |  Conteúdo: --light #f5f5f5
# Cards: branco  |  Primária: #e63428 (vermelho)
# ─────────────────────────────────────────────────────────────────────────────
COR_BG         = "#f0f2f5"   # --light — fundo geral (conteúdo)
COR_PANEL      = "#e8eaed"   # superfície secundária levemente mais escura
COR_CARD       = "#ffffff"   # cards brancos
COR_BORDER     = "#dee2e6"   # --border
COR_ACCENT     = "#e63428"   # --primary (vermelho Abbas)
COR_ACCENT2    = "#c0281d"   # --primary-dk
COR_SUCCESS    = "#198754"   # --success
COR_WARNING    = "#e6a817"   # --warning escurecido p/ leitura em claro
COR_ERROR      = "#dc3545"   # --danger
COR_TEXT       = "#1a1a2e"   # --dark — texto principal
COR_TEXT_DIM   = "#6c757d"   # --gray
COR_HOVER      = "#fef0ef"   # vermelho bem suave p/ hover de linha
COR_XML_OK     = "#d4edda"   # verde claro — XML baixado
COR_HEADER_BG  = "#1a1a2e"   # --dark — navbar
COR_HEADER_FG  = "#ffffff"   # texto no header

FONT_TITLE  = ("Segoe UI", 14, "bold")
FONT_HEADER = ("Segoe UI", 10, "bold")
FONT_BODY   = ("Segoe UI", 10)
FONT_SMALL  = ("Segoe UI", 9)
FONT_MONO   = ("Consolas", 9)


class NFEConsultaApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ABBAS Tecnologia · NF-e Destinadas — SEFAZ Ambiente Nacional")
        self.geometry("1200x780")
        self.minsize(1000, 680)
        self.configure(bg=COR_BG)

        self.cert = None
        self.storage = StorageNSU()
        self._build_ui()
        self._atualizar_status_cert()
        self._carregar_nfes_salvas()
        self._restaurar_cert_salvo()  # restaura certificado da sessão anterior

    # ─── UI ──────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Cabeçalho — igual à .navbar do portal Abbas ───────────────────
        header = tk.Frame(self, bg=COR_HEADER_BG, pady=0)
        header.pack(fill="x", padx=0, pady=0)

        topo = tk.Frame(header, bg=COR_HEADER_BG)
        topo.pack(fill="x", padx=20, pady=(13, 13))

        # border-left 2px vermelho + "ABBAS" branco (nav-title do portal)
        tk.Frame(topo, bg=COR_ACCENT, width=3).pack(side="left", fill="y", padx=(0, 10))

        tk.Label(topo, text="ABBAS", font=("Segoe UI", 14, "bold"),
                 fg=COR_HEADER_FG, bg=COR_HEADER_BG).pack(side="left")
        tk.Label(topo, text=" Tecnologia", font=("Segoe UI", 11),
                 fg="#cccccc", bg=COR_HEADER_BG).pack(side="left")
        tk.Label(topo, text="    NF-e Destinadas  ·  Modelo 55  ·  SEFAZ Ambiente Nacional",
                 font=("Segoe UI", 9), fg="#8899aa", bg=COR_HEADER_BG).pack(side="left", padx=(12, 0))

        # Linha vermelha 3px — brand divider
        tk.Frame(self, bg=COR_ACCENT, height=3).pack(fill="x")
        # ── Corpo principal ────────────────────────────────────────────────
        body = tk.Frame(self, bg=COR_BG)
        body.pack(fill="both", expand=True, padx=20, pady=14)

        # Coluna esquerda (controles)
        left = tk.Frame(body, bg=COR_BG, width=340)
        left.pack(side="left", fill="y", padx=(0, 14))
        left.pack_propagate(False)

        self._build_panel_cert(left)
        self._build_panel_consulta(left)
        self._build_panel_nsu(left)

        # Coluna direita (resultados)
        right = tk.Frame(body, bg=COR_BG)
        right.pack(side="left", fill="both", expand=True)

        self._build_panel_resultados(right)
        self._build_panel_log(right)

        # ── Barra de status ────────────────────────────────────────────────
        self._build_status_bar()

    def _card(self, parent, titulo):
        """Cria um card estilizado com título — tema Abbas."""
        outer = tk.Frame(parent, bg=COR_BORDER, pady=1, padx=1)
        outer.pack(fill="x", pady=(0, 10))
        inner = tk.Frame(outer, bg=COR_CARD, padx=14, pady=12)
        inner.pack(fill="both", expand=True)
        # Linha laranja fina no topo do card
        tk.Frame(inner, bg=COR_ACCENT, height=2).pack(fill="x", pady=(0, 8))
        tk.Label(inner, text=titulo, font=FONT_HEADER,
                 fg=COR_ACCENT, bg=COR_CARD).pack(anchor="w", pady=(0, 6))
        return inner

    def _build_panel_cert(self, parent):
        card = self._card(parent, "CERTIFICADO DIGITAL")

        # Arquivo
        tk.Label(card, text="Arquivo (.pfx / .p12):", font=FONT_SMALL,
                 fg=COR_TEXT_DIM, bg=COR_CARD).pack(anchor="w")

        frow = tk.Frame(card, bg=COR_CARD)
        frow.pack(fill="x", pady=(2, 8))

        self.var_cert_path = tk.StringVar(value="Nenhum certificado carregado")
        self.lbl_cert_path = tk.Label(frow, textvariable=self.var_cert_path,
                                       font=FONT_SMALL, fg=COR_TEXT_DIM,
                                       bg=COR_CARD, anchor="w",
                                       wraplength=220, justify="left")
        self.lbl_cert_path.pack(side="left", fill="x", expand=True)

        tk.Button(frow, text="📁", font=("Consolas", 12),
                  bg=COR_PANEL, fg=COR_ACCENT, relief="flat",
                  activebackground=COR_HOVER, cursor="hand2",
                  command=self._selecionar_cert).pack(side="right")

        # Senha
        tk.Label(card, text="Senha do certificado:", font=FONT_SMALL,
                 fg=COR_TEXT_DIM, bg=COR_CARD).pack(anchor="w")

        self.var_senha = tk.StringVar()
        senha_entry = tk.Entry(card, textvariable=self.var_senha,
                               show="●", font=FONT_BODY,
                               bg=COR_PANEL, fg=COR_TEXT,
                               insertbackground=COR_ACCENT,
                               relief="flat", bd=0, highlightthickness=1,
                               highlightbackground=COR_BORDER,
                               highlightcolor=COR_ACCENT)
        senha_entry.pack(fill="x", pady=(2, 10))
        senha_entry.bind("<Return>", lambda e: self._carregar_cert())

        # Botão carregar
        tk.Button(card, text="Carregar certificado",
                  font=FONT_HEADER, bg=COR_ACCENT, fg="#FFFFFF",
                  activebackground=COR_ACCENT2, activeforeground="#FFFFFF",
                  relief="flat", cursor="hand2", pady=6,
                  command=self._carregar_cert).pack(fill="x")

        # Separador
        tk.Frame(card, bg=COR_BORDER, height=1).pack(fill="x", pady=(12, 8))

        # Botão Windows Store
        import sys as _sys
        if _sys.platform == "win32":
            tk.Button(card, text="Usar certificado instalado no Windows",
                      font=FONT_SMALL, bg=COR_PANEL, fg=COR_ACCENT,
                      activebackground=COR_HOVER, relief="flat",
                      cursor="hand2", pady=5,
                      command=self._selecionar_cert_windows).pack(fill="x")

        # Status do certificado
        self.lbl_cert_status = tk.Label(card, text="",
                                         font=FONT_SMALL, bg=COR_CARD)
        self.lbl_cert_status.pack(anchor="w", pady=(8, 0))

    def _build_panel_consulta(self, parent):
        card = self._card(parent, "CONSULTA SEFAZ")

        # CNPJ
        tk.Label(card, text="CNPJ (somente números):", font=FONT_SMALL,
                 fg=COR_TEXT_DIM, bg=COR_CARD).pack(anchor="w")
        self.var_cnpj = tk.StringVar()
        self.entry_cnpj = tk.Entry(card, textvariable=self.var_cnpj,
                                    font=FONT_BODY, bg=COR_PANEL, fg=COR_TEXT,
                                    insertbackground=COR_ACCENT,
                                    relief="flat", bd=0, highlightthickness=1,
                                    highlightbackground=COR_BORDER,
                                    highlightcolor=COR_ACCENT)
        self.entry_cnpj.pack(fill="x", pady=(2, 10))

        # UF Autor
        uf_row = tk.Frame(card, bg=COR_CARD)
        uf_row.pack(fill="x", pady=(0, 10))
        tk.Label(uf_row, text="Cód. UF (IBGE):", font=FONT_SMALL,
                 fg=COR_TEXT_DIM, bg=COR_CARD).pack(side="left")
        self.var_uf = tk.StringVar(value="43")
        uf_entry = tk.Entry(uf_row, textvariable=self.var_uf, width=5,
                            font=FONT_BODY, bg=COR_PANEL, fg=COR_TEXT,
                            insertbackground=COR_ACCENT, relief="flat", bd=0,
                            highlightthickness=1, highlightbackground=COR_BORDER,
                            highlightcolor=COR_ACCENT)
        uf_entry.pack(side="left", padx=(6, 4))
        tk.Label(uf_row, text="(43=RS  35=SP  41=PR  33=RJ...)", font=FONT_SMALL,
                 fg=COR_TEXT_DIM, bg=COR_CARD).pack(side="left")

        # Ambiente
        tk.Label(card, text="Ambiente:", font=FONT_SMALL,
                 fg=COR_TEXT_DIM, bg=COR_CARD).pack(anchor="w")
        self.var_ambiente = tk.StringVar(value="1")
        amb_frame = tk.Frame(card, bg=COR_CARD)
        amb_frame.pack(anchor="w", pady=(2, 10))
        for txt, val in [("Produção", "1"), ("Homologação", "2")]:
            tk.Radiobutton(amb_frame, text=txt, variable=self.var_ambiente,
                           value=val, font=FONT_SMALL,
                           bg=COR_CARD, fg=COR_TEXT,
                           selectcolor=COR_PANEL,
                           activebackground=COR_CARD,
                           activeforeground=COR_ACCENT).pack(side="left", padx=(0, 14))

        # Botão Consultar
        self.btn_consultar = tk.Button(card, text="Consultar NF-e",
                                        font=FONT_HEADER,
                                        bg=COR_SUCCESS, fg="#FFFFFF",
                                        activebackground="#1E5233",
                                        activeforeground="#FFFFFF",
                                        relief="flat", cursor="hand2", pady=7,
                                        command=self._iniciar_consulta)
        self.btn_consultar.pack(fill="x")

        self.lbl_timer = tk.Label(card, text="", font=FONT_SMALL,
                                   fg=COR_WARNING, bg=COR_CARD)
        self.lbl_timer.pack(anchor="w", pady=(6, 0))

    def _build_panel_nsu(self, parent):
        card = self._card(parent, "CONTROLE NSU")

        self.lbl_nsu_info = tk.Label(card, text="Último NSU: não consultado",
                                      font=FONT_SMALL, fg=COR_TEXT_DIM, bg=COR_CARD,
                                      justify="left")
        self.lbl_nsu_info.pack(anchor="w")

        tk.Label(card, text="\nForçar NSU inicial:", font=FONT_SMALL,
                 fg=COR_TEXT_DIM, bg=COR_CARD).pack(anchor="w")
        self.var_nsu_force = tk.StringVar(value="0")
        tk.Entry(card, textvariable=self.var_nsu_force,
                 font=FONT_BODY, bg=COR_PANEL, fg=COR_TEXT,
                 insertbackground=COR_ACCENT, relief="flat", bd=0,
                 highlightthickness=1, highlightbackground=COR_BORDER,
                 highlightcolor=COR_ACCENT).pack(fill="x", pady=(2, 8))

        tk.Button(card, text="↺  Resetar NSU para este valor",
                  font=FONT_SMALL, bg=COR_PANEL, fg=COR_TEXT_DIM,
                  relief="flat", cursor="hand2",
                  command=self._resetar_nsu).pack(fill="x")

    def _build_panel_resultados(self, parent):
        card_outer = tk.Frame(parent, bg=COR_BORDER, pady=1, padx=1)
        card_outer.pack(fill="both", expand=True, pady=(0, 8))
        card = tk.Frame(card_outer, bg=COR_CARD, padx=14, pady=12)
        card.pack(fill="both", expand=True)

        # Cabeçalho resultados
        hrow = tk.Frame(card, bg=COR_CARD)
        hrow.pack(fill="x", pady=(0, 8))
        tk.Label(hrow, text="NOTAS FISCAIS ENCONTRADAS", font=FONT_HEADER,
                 fg=COR_ACCENT, bg=COR_CARD).pack(side="left")
        self.lbl_total_nfe = tk.Label(hrow, text="", font=FONT_SMALL,
                                       fg=COR_TEXT_DIM, bg=COR_CARD)
        self.lbl_total_nfe.pack(side="right")

        # Treeview
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("NF.Treeview",
                         background=COR_CARD,
                         foreground=COR_TEXT,
                         fieldbackground=COR_CARD,
                         borderwidth=0,
                         font=FONT_MONO,
                         rowheight=24)
        style.configure("NF.Treeview.Heading",
                         background=COR_HEADER_BG,
                         foreground="#ffffff",
                         font=FONT_HEADER,
                         borderwidth=0)
        style.map("NF.Treeview",
                  background=[("selected", "#fef0ef")],
                  foreground=[("selected", COR_ACCENT)])

        cols = ("sel", "nsu", "chave", "emitente", "valor", "emissao", "situacao", "xml")
        self.tree = ttk.Treeview(card, columns=cols, show="headings",
                                  style="NF.Treeview")

        headers = {
            "sel":      ("✓",              30),
            "nsu":      ("NSU",            80),
            "chave":    ("Chave de Acesso", 330),
            "emitente": ("Emitente / CNPJ", 180),
            "valor":    ("Valor (R$)",      100),
            "emissao":  ("Emissão",         120),
            "situacao": ("Situação",         90),
            "xml":      ("XML",              45),
        }
        for col, (label, width) in headers.items():
            self.tree.heading(col, text=label,
                              command=lambda c=col: self._toggle_sel_all() if c == "sel" else None)
            anchor = "center" if col in ("sel", "xml") else "w"
            self.tree.column(col, width=width, anchor=anchor, minwidth=width)

        scroll_y = ttk.Scrollbar(card, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(card, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self._detalhar_nfe)
        self.tree.bind("<Button-1>", self._toggle_sel_item)

        # Botões de ação — linha 1
        brow1 = tk.Frame(card, bg=COR_CARD)
        brow1.pack(fill="x", pady=(8, 3))

        tk.Button(brow1, text="Detalhes", font=FONT_SMALL,
                  bg=COR_PANEL, fg=COR_TEXT, relief="flat",
                  cursor="hand2", padx=10, pady=4,
                  command=lambda: self._detalhar_nfe(None)).pack(side="left", padx=(0, 4))

        tk.Button(brow1, text="⬇ Baixar XML", font=FONT_SMALL,
                  bg=COR_ACCENT, fg="#FFFFFF", activebackground=COR_ACCENT2,
                  relief="flat", cursor="hand2", padx=10, pady=4,
                  command=self._baixar_xml).pack(side="left", padx=(0, 4))

        tk.Button(brow1, text="⬇ Baixar Selecionados", font=FONT_SMALL,
                  bg=COR_ACCENT2, fg="#FFFFFF", activebackground="#B04010",
                  relief="flat", cursor="hand2", padx=10, pady=4,
                  command=self._baixar_xml_lote).pack(side="left", padx=(0, 4))

        tk.Button(brow1, text="Exportar JSON", font=FONT_SMALL,
                  bg=COR_PANEL, fg=COR_TEXT, relief="flat",
                  cursor="hand2", padx=10, pady=4,
                  command=self._exportar_json).pack(side="left", padx=(0, 4))

        tk.Button(brow1, text="Limpar Lista", font=FONT_SMALL,
                  bg=COR_PANEL, fg=COR_ERROR, relief="flat",
                  cursor="hand2", padx=10, pady=4,
                  command=self._limpar_lista).pack(side="right")

        # Botões de ação — linha 2: Manifestação do Destinatário
        brow2 = tk.Frame(card, bg=COR_CARD)
        brow2.pack(fill="x", pady=(0, 2))

        tk.Label(brow2, text="Manifestar:", font=FONT_SMALL,
                 fg=COR_TEXT_DIM, bg=COR_CARD).pack(side="left", padx=(0, 6))

        btn_specs = [
            ("Ciência",         "210210", COR_PANEL,  COR_TEXT),
            ("Confirmação",     "210200", COR_SUCCESS, "#FFFFFF"),
            ("Desconhecimento", "210220", COR_WARNING, "#000000"),
            ("Não Realizada",   "210240", COR_ERROR,   "#FFFFFF"),
        ]
        for label, tp, bg, fg in btn_specs:
            tp_local = tp
            tk.Button(brow2, text=label, font=FONT_SMALL,
                      bg=bg, fg=fg, relief="flat",
                      cursor="hand2", padx=8, pady=3,
                      command=lambda t=tp_local: self._manifestar(t)
                      ).pack(side="left", padx=(0, 3))

    def _build_panel_log(self, parent):
        card_outer = tk.Frame(parent, bg=COR_BORDER, pady=1, padx=1)
        card_outer.pack(fill="x", pady=(0, 0))
        card = tk.Frame(card_outer, bg=COR_CARD, padx=14, pady=8)
        card.pack(fill="both")

        hrow = tk.Frame(card, bg=COR_CARD)
        hrow.pack(fill="x", pady=(0, 4))
        tk.Label(hrow, text="LOG DE EVENTOS", font=FONT_HEADER,
                 fg=COR_ACCENT, bg=COR_CARD).pack(side="left")
        tk.Button(hrow, text="Limpar", font=FONT_SMALL,
                  bg=COR_PANEL, fg=COR_TEXT_DIM, relief="flat",
                  cursor="hand2", command=self._limpar_log).pack(side="right")

        self.log_text = scrolledtext.ScrolledText(
            card, height=7, font=FONT_MONO,
            bg=COR_PANEL, fg=COR_TEXT,
            insertbackground=COR_ACCENT,
            relief="flat", bd=0,
            state="disabled"
        )
        self.log_text.pack(fill="x")
        self.log_text.tag_configure("ok",    foreground=COR_SUCCESS)
        self.log_text.tag_configure("erro",  foreground=COR_ERROR)
        self.log_text.tag_configure("aviso", foreground=COR_WARNING)
        self.log_text.tag_configure("info",  foreground=COR_ACCENT)
        self.log_text.tag_configure("ts",    foreground=COR_TEXT_DIM)

    def _build_status_bar(self):
        bar = tk.Frame(self, bg=COR_BORDER, height=1)
        bar.pack(fill="x")
        status = tk.Frame(self, bg=COR_HEADER_BG, pady=5)
        status.pack(fill="x", padx=0)
        self.lbl_status = tk.Label(status, text="Aguardando certificado...",
                                    font=FONT_SMALL, fg="#8899aa", bg=COR_HEADER_BG,
                                    padx=16)
        self.lbl_status.pack(side="left")
        self.progress = ttk.Progressbar(status, mode="indeterminate", length=120)
        self.progress.pack(side="right", padx=16)

    # ─── AÇÕES ───────────────────────────────────────────────────────────────

    def _selecionar_cert(self):
        path = filedialog.askopenfilename(
            title="Selecionar Certificado Digital",
            filetypes=[("Certificado Digital", "*.pfx *.p12"), ("Todos", "*.*")]
        )
        if path:
            self.var_cert_path.set(os.path.basename(path))
            self._cert_path_full = path
            self._log(f"Arquivo selecionado: {path}", "info")

    def _carregar_cert(self):
        path = getattr(self, "_cert_path_full", None)
        if not path:
            self._log("Selecione um arquivo de certificado primeiro.", "aviso")
            return
        senha = self.var_senha.get()
        if not senha:
            self._log("Informe a senha do certificado.", "aviso")
            return
        try:
            self.cert = CertificadoDigital(path, senha)
            info = self.cert.info()
            self._atualizar_status_cert(ok=True, info=info)
            self._log(f"Certificado carregado: {info['titular']} | CNPJ: {info['cnpj']} | Validade: {info['validade']}", "ok")
            if info["cnpj"]:
                self.var_cnpj.set(info["cnpj"])
            # Persiste configuração (caminho + CNPJ; senha não é salva por segurança)
            self.storage.salvar_cert_config("pfx", path, info["cnpj"])
        except Exception as e:
            self.cert = None
            self._atualizar_status_cert(ok=False)
            self._log(f"Erro ao carregar certificado: {e}", "erro")

    def _restaurar_cert_salvo(self):
        """Tenta restaurar o certificado usado na última sessão."""
        cfg = self.storage.get_cert_config()
        if not cfg:
            return
        tipo  = cfg.get("tipo", "")
        valor = cfg.get("valor", "")
        cnpj  = cfg.get("cnpj", "")

        if tipo == "pfx" and valor and os.path.exists(valor):
            self.var_cert_path.set(os.path.basename(valor))
            self._cert_path_full = valor
            if cnpj:
                self.var_cnpj.set(cnpj)
            self._log(f"Certificado anterior encontrado: {os.path.basename(valor)} — "
                      f"informe a senha e clique em Carregar.", "info")

        elif tipo == "windows" and valor:
            # Tenta recarregar do Windows Store pelo CN salvo
            certs = listar_certificados_windows()
            for c in certs:
                if c["titular"] == valor:
                    self._log(f"Restaurando certificado do Windows: {valor}...", "info")
                    self._carregar_cert_do_windows(c)
                    return
            self._log(f"Certificado Windows '{valor}' não encontrado no store.", "aviso")

    def _selecionar_cert_windows(self):
        """Abre janela de seleção de certificados instalados no Windows."""
        self._log("Lendo certificados instalados no Windows...", "info")
        certs = listar_certificados_windows()

        if not certs:
            messagebox.showwarning(
                "Nenhum certificado encontrado",
                "Não foram encontrados certificados válidos com chave privada\n"
                "no repositório Pessoal (MY) do Windows.\n\n"
                "Verifique se o certificado A1 está instalado em:\n"
                "Gerenciador de Certificados → Certificados - Usuário Atual → Pessoal"
            )
            return

        # Janela de seleção
        win = tk.Toplevel(self)
        win.title("Certificados instalados no Windows")
        win.geometry("620x340")
        win.configure(bg=COR_BG)
        win.grab_set()

        tk.Label(win, text="Selecione o certificado:", font=FONT_HEADER,
                 fg=COR_TEXT, bg=COR_BG).pack(padx=20, pady=(16, 8), anchor="w")

        # Lista de certificados
        frame_list = tk.Frame(win, bg=COR_BORDER, padx=1, pady=1)
        frame_list.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        lb = tk.Listbox(frame_list, font=FONT_BODY, bg=COR_CARD, fg=COR_TEXT,
                        selectbackground=COR_ACCENT, selectforeground="#FFFFFF",
                        relief="flat", bd=0, activestyle="none")
        lb.pack(fill="both", expand=True)

        for c in certs:
            cnpj_fmt = c["cnpj"] if c["cnpj"] else "CPF/CNPJ não extraído"
            lb.insert("end", f"  {c['titular']}   |   {cnpj_fmt}   |   Val: {c['validade']}")

        if certs:
            lb.selection_set(0)

        def confirmar():
            sel = lb.curselection()
            if not sel:
                return
            cert_escolhido = certs[sel[0]]
            win.destroy()
            self._carregar_cert_do_windows(cert_escolhido)

        btn_row = tk.Frame(win, bg=COR_BG)
        btn_row.pack(fill="x", padx=20, pady=(0, 16))
        tk.Button(btn_row, text="Usar este certificado", font=FONT_HEADER,
                  bg=COR_ACCENT, fg="#FFFFFF", relief="flat",
                  cursor="hand2", pady=6, command=confirmar).pack(side="left")
        tk.Button(btn_row, text="Cancelar", font=FONT_SMALL,
                  bg=COR_PANEL, fg=COR_TEXT_DIM, relief="flat",
                  cursor="hand2", pady=6, padx=12,
                  command=win.destroy).pack(side="left", padx=(8, 0))

        lb.bind("<Double-1>", lambda e: confirmar())

    def _carregar_cert_do_windows(self, cert_info: dict):
        """Carrega certificado do Windows Store. Tenta exportar; se falhar, usa WinHTTP."""
        titular = cert_info["titular"]
        self._log(f"Carregando certificado do Windows Store: {titular}...", "info")

        # Tenta primeiro exportar (para certificados exportáveis)
        try:
            from src.win_cert_store import exportar_pem_do_store
            cert_pem, key_pem = exportar_pem_do_store(cert_info["der"])
            self.cert = _CertificadoWindowsStore(cert_pem, key_pem, cert_info)
            info = self.cert.info()
            self._atualizar_status_cert(ok=True, info=info)
            self._log(
                f"Certificado exportável carregado: {info['titular']} | "
                f"CNPJ: {info['cnpj']} | Validade: {info['validade']}", "ok"
            )
            if info["cnpj"]:
                self.var_cnpj.set(info["cnpj"])
            return

        except Exception as e_export:
            self._log(f"Exportação bloqueada ({e_export.__class__.__name__}) — "
                      f"usando modo WinHTTP (não-exportável)...", "aviso")

        # Fallback: modo WinHTTP — chave nunca sai do Windows
        try:
            import win32com.client  # Verifica se pywin32 está instalado
            self.cert = _CertificadoWinHTTP(cert_info)
            info = self.cert.info()
            self._atualizar_status_cert(ok=True, info=info)
            self._log(
                f"Certificado não-exportável carregado via WinHTTP: {info['titular']} | "
                f"CNPJ: {info['cnpj']} | Validade: {info['validade']}", "ok"
            )
            if info["cnpj"]:
                self.var_cnpj.set(info["cnpj"])
            self.storage.salvar_cert_config("windows", cert_info["titular"], info["cnpj"])

        except ImportError:
            self.cert = None
            self._atualizar_status_cert(ok=False)
            self._log(
                "pywin32 não instalado — necessário para certificados não-exportáveis.\n"
                "Execute: pip install pywin32", "erro"
            )
            messagebox.showerror(
                "pywin32 necessário",
                "Este certificado foi instalado como NÃO-EXPORTÁVEL.\n\n"
                "Para usá-lo sem o arquivo .pfx, instale o pywin32:\n\n"
                "    pip install pywin32\n\n"
                "Depois reinicie a aplicação."
            )
        except Exception as e:
            self.cert = None
            self._atualizar_status_cert(ok=False)
            self._log(f"Erro ao carregar via WinHTTP: {e}", "erro")

    def _atualizar_status_cert(self, ok=None, info=None):
        if ok is True and info:
            txt = f"✔ {info['titular']}  |  Val: {info['validade']}"
            self.lbl_cert_status.config(text=txt, fg=COR_SUCCESS)
            self.lbl_status.config(text=f"Certificado carregado · {info['titular']}")
        elif ok is False:
            self.lbl_cert_status.config(text="✘ Falha ao carregar certificado", fg=COR_ERROR)
            self.lbl_status.config(text="Erro no certificado")
        else:
            self.lbl_cert_status.config(text="Nenhum certificado carregado", fg=COR_TEXT_DIM)

    def _iniciar_consulta(self):
        if not self.cert:
            messagebox.showwarning("Atenção", "Carregue o certificado digital antes de consultar.")
            return
        cnpj = self.var_cnpj.get().strip().replace(".", "").replace("/", "").replace("-", "")
        if len(cnpj) != 14 or not cnpj.isdigit():
            messagebox.showwarning("Atenção", "CNPJ inválido. Informe somente 14 dígitos.")
            return

        self.btn_consultar.config(state="disabled")
        self.progress.start(10)
        self._log(f"Iniciando consulta para CNPJ {cnpj}...", "info")
        self._set_status("Consultando SEFAZ...")

        thread = threading.Thread(target=self._executar_consulta,
                                  args=(cnpj,), daemon=True)
        thread.start()

    def _executar_consulta(self, cnpj):
        try:
            ambiente  = int(self.var_ambiente.get())
            cuf_autor = self.var_uf.get().strip() or "43"
            client    = SefazClient(self.cert, ambiente, cuf_autor)

            total_notas  = 0
            rodada       = 0
            ultimo_nsu   = self.storage.get_nsu(cnpj)

            self._log(f"Último NSU registrado: {ultimo_nsu:015d}", "info")

            # Loop: continua consultando enquanto houver NSUs novos
            while True:
                rodada += 1
                self._set_status(f"Consultando SEFAZ... rodada {rodada} (NSU {ultimo_nsu:015d})")
                resultado = client.consultar_distribuicao(cnpj, ultimo_nsu)

                if resultado["status"] == "erro":
                    self._log(f"Erro SEFAZ [{resultado['codigo']}]: {resultado['mensagem']}", "erro")
                    self._set_status(f"Erro: {resultado['mensagem']}")
                    return

                if resultado["status"] == "consumo_indevido":
                    self._log("Consumo indevido — aguarde 60 minutos antes de nova consulta.", "aviso")
                    self._set_status("Aguardando cooldown de 60 minutos")
                    self.after(0, self._iniciar_countdown)
                    return

                codigo   = resultado.get("codigo", "?")
                mensagem = resultado.get("mensagem", "")
                nfes     = resultado.get("notas", [])
                novo_nsu = resultado.get("max_nsu", ultimo_nsu)

                self.storage.set_nsu(cnpj, novo_nsu)
                if nfes:
                    self.storage.salvar_notas(cnpj, nfes)
                    total_notas += len(nfes)
                    self.after(0, self._atualizar_resultados, nfes, novo_nsu, cnpj)

                self._log(
                    f"Rodada {rodada} [{codigo}] {mensagem} — "
                    f"{len(nfes)} doc(s) · NSU atual: {novo_nsu:015d}", "info"
                )

                # Para quando não há mais documentos novos
                if resultado["status"] == "sem_novos" or novo_nsu <= ultimo_nsu:
                    break

                # Para quando o SEFAZ diz que chegamos ao maxNSU
                max_nsu_resp = resultado.get("max_nsu", 0)
                if novo_nsu >= max_nsu_resp and resultado["status"] != "ok":
                    break

                ultimo_nsu = novo_nsu

                # Se veio menos que 50 documentos, provavelmente esgotou
                if len(nfes) < 50:
                    # Faz mais uma consulta para confirmar que não há mais
                    resultado2 = client.consultar_distribuicao(cnpj, novo_nsu)
                    if resultado2["status"] in ("sem_novos", "erro"):
                        codigo2 = resultado2.get("codigo", "?")
                        self._log(f"Confirmado fim do lote [{codigo2}]: {resultado2.get('mensagem','')}", "info")
                        break
                    elif resultado2["status"] == "consumo_indevido":
                        self._log("Limite atingido — aguarde 60 minutos.", "aviso")
                        self.after(0, self._iniciar_countdown)
                        return
                    else:
                        # Ainda há mais — continua no próximo loop
                        nfes2    = resultado2.get("notas", [])
                        novo_nsu = resultado2.get("max_nsu", novo_nsu)
                        self.storage.set_nsu(cnpj, novo_nsu)
                        if nfes2:
                            self.storage.salvar_notas(cnpj, nfes2)
                            total_notas += len(nfes2)
                            self.after(0, self._atualizar_resultados, nfes2, novo_nsu, cnpj)
                        rodada += 1
                        self._log(
                            f"Rodada {rodada} [{resultado2.get('codigo','?')}] — "
                            f"{len(nfes2)} doc(s) · NSU: {novo_nsu:015d}", "info"
                        )
                        ultimo_nsu = novo_nsu
                        if len(nfes2) < 50:
                            break

            self._log(
                f"Consulta encerrada — {total_notas} documento(s) recebido(s) em {rodada} rodada(s). "
                f"NSU final: {ultimo_nsu:015d}", "ok"
            )
            self._set_status(f"Concluído · {total_notas} doc(s) · NSU: {ultimo_nsu:015d}")

        except Exception as e:
            self._log(f"Erro inesperado: {e}", "erro")
            self._set_status(f"Erro: {e}")
        finally:
            self.after(0, self._finalizar_consulta)

    def _finalizar_consulta(self):
        self.progress.stop()
        self.btn_consultar.config(state="normal")
        self._atualizar_nsu_label()

    def _atualizar_resultados(self, nfes, novo_nsu, cnpj):
        for nfe in nfes:
            self._inserir_nfe_tree(nfe)
        total = len(self.tree.get_children())
        self.lbl_total_nfe.config(text=f"{total} nota(s) na lista")
        self._atualizar_nsu_label()

    def _inserir_nfe_tree(self, nfe):
        nsu_str = str(nfe.get("nsu", "")).strip()
        chave   = nfe.get("chave", "").strip()
        tipo    = nfe.get("tipo", "")

        # Deduplica por NSU
        for item in self.tree.get_children():
            if self.tree.item(item, "values")[1].strip() == nsu_str:
                return

        sit = nfe.get("situacao", "")
        xml_baixado = bool(nfe.get("xml_raw", ""))

        if sit == "Autorizada":
            cor = "ok"
        elif sit in ("Cancelada", "Denegada", "Cancelamento"):
            cor = "cancel"
        elif "Evento" in tipo:
            cor = "evento"
        else:
            cor = ""

        emitente = nfe.get("emitente", "")
        if "Evento" in tipo and not emitente:
            emitente = tipo

        self.tree.insert("", "end", values=(
            "",           # sel (checkbox visual)
            nsu_str,
            chave,
            emitente,
            nfe.get("valor", ""),
            nfe.get("emissao", ""),
            sit,
            "✔" if xml_baixado else "",
        ), tags=(cor,))

        self.tree.tag_configure("ok",     foreground=COR_SUCCESS)
        self.tree.tag_configure("cancel", foreground=COR_ERROR)
        self.tree.tag_configure("evento", foreground=COR_WARNING)
        self.tree.tag_configure("xml_ok", background=COR_XML_OK)

    def _toggle_sel_all(self):
        """Marca/desmarca todas as linhas."""
        children = self.tree.get_children()
        if not children:
            return
        # Verifica se alguma está desmarcada
        alguma_desmarcada = any(
            self.tree.item(i, "values")[0] == "" for i in children
        )
        novo = "☑" if alguma_desmarcada else ""
        for item in children:
            vals = list(self.tree.item(item, "values"))
            vals[0] = novo
            self.tree.item(item, values=vals)

    def _toggle_sel_item(self, event):
        """Clique na coluna ✓ alterna seleção individual."""
        region = self.tree.identify_region(event.x, event.y)
        col    = self.tree.identify_column(event.x)
        if region == "cell" and col == "#1":  # coluna sel
            item = self.tree.identify_row(event.y)
            if item:
                vals = list(self.tree.item(item, "values"))
                vals[0] = "" if vals[0] == "☑" else "☑"
                self.tree.item(item, values=vals)
            return "break"

    def _get_selecionados(self):
        """Retorna lista de dicts das notas marcadas com ☑."""
        cnpj  = self.var_cnpj.get().strip().replace(".", "").replace("/", "").replace("-", "")
        notas = []
        for item in self.tree.get_children():
            vals = self.tree.item(item, "values")
            if vals[0] == "☑":
                nsu_str = vals[1].strip()
                chave   = vals[2].strip()
                nota    = self.storage.get_nota(cnpj, chave) or {"nsu": int(nsu_str), "chave": chave}
                notas.append(nota)
        return notas

    def _atualizar_nsu_label(self):
        cnpj = self.var_cnpj.get().strip().replace(".", "").replace("/", "").replace("-", "")
        if cnpj and cnpj.isdigit() and len(cnpj) == 14:
            nsu = self.storage.get_nsu(cnpj)
            self.lbl_nsu_info.config(
                text=f"CNPJ: {cnpj}\nÚltimo NSU: {nsu:015d}"
            )

    def _resetar_nsu(self):
        cnpj = self.var_cnpj.get().strip().replace(".", "").replace("/", "").replace("-", "")
        if not cnpj or not cnpj.isdigit() or len(cnpj) != 14:
            messagebox.showwarning("Atenção", "CNPJ inválido.")
            return
        try:
            nsu = int(self.var_nsu_force.get())
            self.storage.set_nsu(cnpj, nsu)
            self._atualizar_nsu_label()
            self._log(f"NSU resetado para {nsu:015d} (CNPJ {cnpj})", "aviso")
        except ValueError:
            messagebox.showwarning("Atenção", "NSU inválido.")

    def _iniciar_countdown(self):
        """Conta regressiva de 60 minutos para nova consulta."""
        segundos_restantes = [3600]

        def tick():
            if segundos_restantes[0] > 0:
                m = segundos_restantes[0] // 60
                s = segundos_restantes[0] % 60
                self.lbl_timer.config(text=f"⏱ Próxima consulta em: {m:02d}:{s:02d}")
                segundos_restantes[0] -= 1
                self.after(1000, tick)
            else:
                self.lbl_timer.config(text="✔ Pronto para nova consulta")
                self.btn_consultar.config(state="normal")

        self.btn_consultar.config(state="disabled")
        tick()

    def _detalhar_nfe(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        cnpj = self.var_cnpj.get().strip().replace(".", "").replace("/", "").replace("-", "")
        chave = vals[1].strip()
        nota = self.storage.get_nota(cnpj, chave)

        win = tk.Toplevel(self)
        win.title(f"Detalhes NF-e  ·  {chave[:10]}...")
        win.geometry("700x520")
        win.configure(bg=COR_BG)
        win.grab_set()

        tk.Label(win, text="Detalhes da NF-e", font=FONT_TITLE,
                 fg=COR_TEXT, bg=COR_BG).pack(padx=20, pady=14, anchor="w")

        txt = scrolledtext.ScrolledText(win, font=FONT_MONO,
                                         bg=COR_PANEL, fg=COR_TEXT,
                                         relief="flat", bd=0)
        txt.pack(fill="both", expand=True, padx=20, pady=(0, 14))
        txt.insert("end", json.dumps(nota, indent=2, ensure_ascii=False) if nota else
                   f"Chave: {chave}\nNSU: {vals[0]}\nEmitente: {vals[2]}\nValor: {vals[3]}\nEmissão: {vals[4]}\nSituação: {vals[5]}")
        txt.config(state="disabled")

    def _baixar_xml(self):
        """Baixa o XML completo da NF-e selecionada via consNSU."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Atenção", "Selecione uma NF-e na lista.")
            return
        if not self.cert:
            messagebox.showwarning("Atenção", "Carregue o certificado digital primeiro.")
            return

        vals  = self.tree.item(sel[0], "values")
        nsu_str = vals[0].strip()
        chave   = vals[1].strip()

        try:
            nsu = int(nsu_str)
        except ValueError:
            messagebox.showerror("Erro", f"NSU inválido: {nsu_str}")
            return

        cnpj = self.var_cnpj.get().strip().replace(".", "").replace("/", "").replace("-", "")

        # Verifica se já temos XML salvo localmente
        nota = self.storage.get_nota(cnpj, chave)
        xml_local = nota.get("xml_raw", "") if nota else ""
        if xml_local:
            self._salvar_xml_dialog(xml_local, chave)
            return

        # Busca no SEFAZ
        self._log(f"Baixando XML · NSU {nsu:015d}...", "info")
        self._set_status("Baixando XML...")
        self.progress.start(10)

        def worker():
            try:
                ambiente  = int(self.var_ambiente.get())
                client    = EventoClient(self.cert, ambiente)
                resultado = client.baixar_xml(cnpj, nsu)

                if resultado["status"] == "ok" and resultado.get("xml_str"):
                    xml_str = resultado["xml_str"]
                    # Salva no storage para não precisar baixar de novo
                    if nota:
                        nota["xml_raw"] = xml_str
                        self.storage.salvar_notas(cnpj, [nota])
                    self._log(f"XML baixado com sucesso · {len(xml_str)} bytes", "ok")
                    self.after(0, self._salvar_xml_dialog, xml_str, chave)
                elif resultado["status"] == "sem_xml":
                    self._log(
                        f"XML completo não disponível · [{resultado['codigo']}] "
                        f"{resultado['mensagem']} · "
                        f"Faça a Manifestação (Ciência) primeiro.", "aviso"
                    )
                    self.after(0, messagebox.showwarning, "XML indisponível",
                               "O XML completo ainda não está disponível.\n\n"
                               "Faça a Manifestação → Ciência da Operação primeiro.\n"
                               "Após a manifestação, o SEFAZ libera o XML na próxima consulta.")
                else:
                    msg = resultado.get("mensagem", "Erro desconhecido")
                    self._log(f"Erro ao baixar XML: {msg}", "erro")
            except Exception as e:
                self._log(f"Erro inesperado ao baixar XML: {e}", "erro")
            finally:
                self.after(0, self.progress.stop)
                self.after(0, lambda: self._set_status("Pronto"))

        threading.Thread(target=worker, daemon=True).start()

    def _salvar_xml_dialog(self, xml_str: str, chave: str):
        """Abre diálogo para salvar o XML em disco."""
        nome_sugerido = f"NFe_{chave[:20]}.xml" if chave else "NFe.xml"
        path = filedialog.asksaveasfilename(
            defaultextension=".xml",
            filetypes=[("XML NF-e", "*.xml"), ("Todos", "*.*")],
            initialfile=nome_sugerido,
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(xml_str)
            self._log(f"XML salvo em: {path}", "ok")

    def _manifestar(self, tp_evento: str):
        """Envia evento de manifestação do destinatário para a NF-e selecionada."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Atenção", "Selecione uma NF-e na lista.")
            return
        if not self.cert:
            messagebox.showwarning("Atenção", "Carregue o certificado digital primeiro.")
            return

        vals  = self.tree.item(sel[0], "values")
        chave = vals[1].strip()
        cnpj  = self.var_cnpj.get().strip().replace(".", "").replace("/", "").replace("-", "")

        desc_evento = TIPOS_EVENTO.get(tp_evento, tp_evento)

        # Para "Operação não Realizada" pede justificativa
        justificativa = ""
        if tp_evento == "210240":
            win = tk.Toplevel(self)
            win.title("Justificativa")
            win.geometry("420x160")
            win.configure(bg=COR_BG)
            win.grab_set()
            tk.Label(win, text="Justificativa (obrigatória, 15–255 caracteres):",
                     font=FONT_SMALL, fg=COR_TEXT_DIM, bg=COR_BG).pack(padx=16, pady=(14, 4), anchor="w")
            var_just = tk.StringVar()
            tk.Entry(win, textvariable=var_just, font=FONT_BODY,
                     bg=COR_PANEL, fg=COR_TEXT, relief="flat",
                     insertbackground=COR_ACCENT).pack(fill="x", padx=16, pady=(0, 10))
            resultado_just = [None]

            def confirmar():
                j = var_just.get().strip()
                if len(j) < 15:
                    messagebox.showwarning("Atenção", "Justificativa deve ter pelo menos 15 caracteres.", parent=win)
                    return
                resultado_just[0] = j
                win.destroy()

            tk.Button(win, text="Confirmar", font=FONT_HEADER,
                      bg=COR_ACCENT, fg="#FFFFFF", relief="flat",
                      command=confirmar).pack(padx=16, fill="x")
            win.wait_window()
            if resultado_just[0] is None:
                return
            justificativa = resultado_just[0]

        # Confirmação
        if not messagebox.askyesno(
            "Confirmar Manifestação",
            f"Evento: {desc_evento}\n\nChave: {chave}\n\n"
            f"⚠ Esta ação será registrada permanentemente no SEFAZ.\n\nConfirmar?"
        ):
            return

        self._log(f"Enviando manifestação [{tp_evento}] {desc_evento} · Chave: {chave[:20]}...", "info")
        self.progress.start(10)

        def worker():
            try:
                ambiente  = int(self.var_ambiente.get())
                client    = EventoClient(self.cert, ambiente)
                resultado = client.manifestar(cnpj, chave, tp_evento, justificativa)

                if resultado["status"] == "ok":
                    prot = resultado.get("protocolo", "")
                    self._log(
                        f"Manifestação registrada com sucesso! "
                        f"Protocolo: {prot} · [{resultado['codigo']}] {resultado['mensagem']}",
                        "ok"
                    )
                    # Atualiza situação na tabela e no storage
                    self.after(0, self._atualizar_situacao_tree, chave, desc_evento)
                    nota = self.storage.get_nota(cnpj, chave)
                    if nota:
                        nota["situacao"] = desc_evento
                        nota["protocolo_manifestacao"] = prot
                        self.storage.salvar_notas(cnpj, [nota])
                else:
                    self._log(
                        f"Erro na manifestação · [{resultado['codigo']}] {resultado['mensagem']}",
                        "erro"
                    )
            except Exception as e:
                self._log(f"Erro inesperado na manifestação: {e}", "erro")
            finally:
                self.after(0, self.progress.stop)
                self.after(0, lambda: self._set_status("Pronto"))

        threading.Thread(target=worker, daemon=True).start()

    def _baixar_xml_lote(self):
        """Baixa XMLs de todas as notas marcadas com ☑, salvando numa pasta."""
        selecionadas = self._get_selecionados()
        if not selecionadas:
            messagebox.showwarning("Atenção", "Marque as notas desejadas clicando na coluna ✓.\n"
                                              "Clique no cabeçalho ✓ para selecionar todas.")
            return
        if not self.cert:
            messagebox.showwarning("Atenção", "Carregue o certificado digital primeiro.")
            return

        pasta = filedialog.askdirectory(title="Selecione a pasta para salvar os XMLs")
        if not pasta:
            return

        cnpj = self.var_cnpj.get().strip().replace(".", "").replace("/", "").replace("-", "")
        total = len(selecionadas)
        self._log(f"Download em lote: {total} nota(s) selecionada(s)...", "info")
        self.progress.start(10)
        self.btn_consultar.config(state="disabled")

        def worker():
            ok = 0
            erros = 0
            ambiente = int(self.var_ambiente.get())
            client   = EventoClient(self.cert, ambiente)

            for nota in selecionadas:
                chave = nota.get("chave", "").strip()
                nsu   = nota.get("nsu", 0)

                if not chave and not nsu:
                    continue

                # Verifica XML já salvo localmente
                xml_str = nota.get("xml_raw", "")

                if not xml_str and nsu:
                    try:
                        res = client.baixar_xml(cnpj, int(nsu))
                        if res.get("status") == "ok":
                            xml_str = res["xml_str"]
                            nota["xml_raw"] = xml_str
                            self.storage.salvar_notas(cnpj, [nota])
                    except Exception as e:
                        self._log(f"Erro ao baixar NSU {nsu}: {e}", "erro")
                        erros += 1
                        continue

                if xml_str:
                    nome = f"NFe_{chave or nsu}.xml"
                    caminho = os.path.join(pasta, nome)
                    with open(caminho, "w", encoding="utf-8") as f:
                        f.write(xml_str)
                    self._log(f"  ✔ {nome}", "ok")
                    self.after(0, self._marcar_xml_baixado, nsu)
                    ok += 1
                else:
                    self._log(f"  ✘ NSU {nsu} — XML não disponível (faça Ciência primeiro)", "aviso")
                    erros += 1

            self._log(f"Download concluído: {ok} salvo(s), {erros} sem XML disponível.", "ok" if not erros else "aviso")
            self.after(0, self.progress.stop)
            self.after(0, lambda: self.btn_consultar.config(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    def _marcar_xml_baixado(self, nsu):
        """Atualiza a coluna XML na linha correspondente ao NSU."""
        nsu_str = str(nsu).strip()
        for item in self.tree.get_children():
            vals = list(self.tree.item(item, "values"))
            if vals[1].strip() == nsu_str:
                vals[7] = "✔"
                self.tree.item(item, values=vals)
                break

    def _atualizar_situacao_tree(self, chave: str, nova_sit: str):
        """Atualiza a coluna Situação na Treeview para a nota especificada."""
        for item in self.tree.get_children():
            vals = list(self.tree.item(item, "values"))
            if vals[2].strip() == chave:
                vals[6] = nova_sit
                self.tree.item(item, values=vals)
                break

    def _exportar_json(self):
        cnpj = self.var_cnpj.get().strip().replace(".", "").replace("/", "").replace("-", "")
        notas = self.storage.get_todas_notas(cnpj)
        if not notas:
            messagebox.showinfo("Exportar", "Nenhuma nota para exportar.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialfile=f"nfe_{cnpj}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(notas, f, indent=2, ensure_ascii=False)
            self._log(f"Exportado: {path}", "ok")

    def _carregar_nfes_salvas(self):
        """Carrega notas previamente consultadas do storage local."""
        cnpj = self.var_cnpj.get().strip()
        notas = self.storage.get_todas_notas(cnpj) if cnpj else []
        for nota in notas:
            self._inserir_nfe_tree(nota)

    def _limpar_lista(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.lbl_total_nfe.config(text="")

    def _limpar_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    def _log(self, msg, tipo="info"):
        self.log_text.config(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{ts}] ", "ts")
        self.log_text.insert("end", f"{msg}\n", tipo)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _set_status(self, msg):
        self.after(0, lambda: self.lbl_status.config(text=msg))




# ─────────────────────────────────────────────────────────────────────────────
class _CertificadoWinHTTP:
    """
    Certificado do Windows Store NÃO-EXPORTÁVEL.
    A chave privada nunca sai do Windows — o SChannel gerencia a autenticação.
    Sinaliza ao SefazClient para usar WinHTTP COM ao invés de requests+PEM.
    """
    def __init__(self, cert_info: dict):
        self._titular  = cert_info.get("titular", "")
        self._cnpj     = cert_info.get("cnpj", "")
        self._validade = cert_info.get("validade", "")
        # Usado pelo SefazClient para detectar o modo WinHTTP
        self._winhttp_cn = self._titular

    def info(self) -> dict:
        return {
            "titular":  self._titular,
            "cnpj":     self._cnpj,
            "validade": self._validade,
            "arquivo":  "Windows Store (não-exportável / WinHTTP)",
        }

    def exportar_pem_temp(self):
        raise RuntimeError(
            "Certificado não-exportável — use SefazClient em modo WinHTTP."
        )

    @property
    def titular(self) -> str:
        return self._titular

    @property
    def cnpj(self) -> str:
        return self._cnpj

    @property
    def validade(self) -> str:
        return self._validade


class _CertificadoWindowsStore:
    """
    Wrapper que expõe a mesma interface que CertificadoDigital,
    mas usa PEMs já extraídos do Windows Certificate Store.
    Não precisa de arquivo .pfx nem de senha.
    """
    def __init__(self, cert_pem: bytes, key_pem: bytes, cert_info: dict):
        self._cert_pem = cert_pem
        self._key_pem  = key_pem
        self._titular  = cert_info.get("titular", "")
        self._cnpj     = cert_info.get("cnpj", "")
        self._validade = cert_info.get("validade", "")

    def info(self) -> dict:
        return {
            "titular":  self._titular,
            "cnpj":     self._cnpj,
            "validade": self._validade,
            "arquivo":  "Windows Certificate Store",
        }

    def pem_cert(self) -> bytes:
        return self._cert_pem

    def pem_key(self) -> bytes:
        return self._key_pem

    def exportar_pem_temp(self):
        """Cria arquivos PEM temporários e retorna (cert_path, key_path)."""
        import tempfile, os
        cert_file = tempfile.NamedTemporaryFile(delete=False, suffix=".cert.pem")
        cert_file.write(self._cert_pem)
        cert_file.close()
        key_file = tempfile.NamedTemporaryFile(delete=False, suffix=".key.pem")
        key_file.write(self._key_pem)
        key_file.close()
        return cert_file.name, key_file.name

    @property
    def titular(self) -> str:
        return self._titular

    @property
    def cnpj(self) -> str:
        return self._cnpj

    @property
    def validade(self) -> str:
        return self._validade


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = NFEConsultaApp()
    app.mainloop()
