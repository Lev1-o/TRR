import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import pandas as pd
import os
import sys
import io
from datetime import datetime

# ============================================
# CONFIGURAÇÃO INICIAL
# ============================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

# ============================================
# CORES
# ============================================
AZUL_ESCURO = "#041E42"
VERDE       = "#00DB81"
CINZA       = "#2b2b2b"
BRANCO      = "#ffffff"
AMARELO     = "#ffc107"
VERMELHO    = "#dc3545"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("EQUIPO | Sistema de Gestão de Revisões")
        self.geometry("1280x780")
        self.minsize(1100, 700)
        self.configure(fg_color="#0d1117")
        self._build_ui()

    # ============================================
    # UI PRINCIPAL
    # ============================================
    def _build_ui(self):
        # ---------- SIDEBAR ----------
        self.sidebar = ctk.CTkFrame(self, width=230, fg_color=AZUL_ESCURO, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        ctk.CTkLabel(
            self.sidebar, text="⚙ EQUIPO",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=VERDE
        ).pack(pady=(30, 5))

        ctk.CTkLabel(
            self.sidebar, text="Gestão de Revisões",
            font=ctk.CTkFont(size=12), text_color="#aaaaaa"
        ).pack(pady=(0, 30))

        self.nav_buttons = {}
        pages = [
            ("🏠  Dashboard",    "dashboard"),
            ("📡  KM Atual",     "km_atual"),
            ("📧  Prospecção",   "prospeccao"),
            ("📋  Histórico",    "historico"),
        ]
        for label, key in pages:
            btn = ctk.CTkButton(
                self.sidebar, text=label, anchor="w", height=45,
                fg_color="transparent", hover_color="#0a2d5e",
                text_color=BRANCO, font=ctk.CTkFont(size=13),
                command=lambda k=key: self.show_page(k)
            )
            btn.pack(fill="x", padx=10, pady=4)
            self.nav_buttons[key] = btn

        ctk.CTkLabel(
            self.sidebar,
            text=f"v1.0.0  |  {datetime.now().strftime('%d/%m/%Y')}",
            font=ctk.CTkFont(size=10), text_color="#555555"
        ).pack(side="bottom", pady=15)

        # ---------- ÁREA PRINCIPAL ----------
        self.main_area = ctk.CTkFrame(self, fg_color="#0d1117", corner_radius=0)
        self.main_area.pack(side="left", fill="both", expand=True)

        self.pages = {
            "dashboard":  self._build_dashboard(),
            "km_atual":   self._build_km_atual(),
            "prospeccao": self._build_prospeccao(),
            "historico":  self._build_historico(),
        }
        self.show_page("dashboard")

    def show_page(self, name):
        for frame in self.pages.values():
            frame.pack_forget()
        self.pages[name].pack(fill="both", expand=True)
        for key, btn in self.nav_buttons.items():
            btn.configure(fg_color=AZUL_ESCURO if key == name else "transparent")

    # ============================================
    # COMPONENTE: CARD
    # ============================================
    def _card(self, parent, title, value, icon, color):
        frame = ctk.CTkFrame(parent, fg_color="#161b22", corner_radius=12, width=190, height=105)
        frame.pack_propagate(False)
        ctk.CTkLabel(frame, text=icon,  font=ctk.CTkFont(size=26)).pack(pady=(12, 0))
        lbl = ctk.CTkLabel(frame, text=value, font=ctk.CTkFont(size=20, weight="bold"), text_color=color)
        lbl.pack()
        lbl._value_label = True
        ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=11), text_color="#888888").pack()
        return frame

    def _set_card_value(self, card, value):
        for w in card.winfo_children():
            if isinstance(w, ctk.CTkLabel) and hasattr(w, "_value_label"):
                w.configure(text=value)
                return

    # ============================================
    # UTILITÁRIOS DE LOG
    # ============================================
    def _log(self, textbox, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        textbox.configure(state="normal")
        textbox.insert("end", f"[{ts}] {msg}\n")
        textbox.see("end")
        textbox.configure(state="disabled")

    def _limpar_log(self, textbox):
        textbox.configure(state="normal")
        textbox.delete("1.0", "end")
        textbox.configure(state="disabled")

    def _make_log_box(self, parent, height=300):
        box = ctk.CTkTextbox(
            parent, height=height,
            fg_color="#161b22", text_color="#00FF88",
            font=ctk.CTkFont(family="Courier New", size=12),
            state="disabled"
        )
        return box

    def _make_progress(self, parent):
        bar = ctk.CTkProgressBar(parent, height=12, fg_color="#21262d", progress_color=VERDE)
        bar.pack(fill="x", padx=30, pady=(0, 4))
        bar.set(0)
        lbl = ctk.CTkLabel(parent, text="Aguardando execução...",
                            font=ctk.CTkFont(size=12), text_color="#888888")
        lbl.pack(anchor="w", padx=30, pady=(0, 12))
        return bar, lbl

    # ============================================
    # DASHBOARD
    # ============================================
    def _build_dashboard(self):
        page = ctk.CTkFrame(self.main_area, fg_color="#0d1117", corner_radius=0)

        ctk.CTkLabel(page, text="🏠  Dashboard",
                     font=ctk.CTkFont(size=24, weight="bold"), text_color=BRANCO
                     ).pack(anchor="w", padx=30, pady=(25, 3))
        ctk.CTkLabel(page, text="Visão geral do sistema de gestão de revisões",
                     font=ctk.CTkFont(size=13), text_color="#888888"
                     ).pack(anchor="w", padx=30, pady=(0, 18))

        # Cards
        row = ctk.CTkFrame(page, fg_color="transparent")
        row.pack(anchor="w", padx=30, pady=(0, 12))
        self.card_total    = self._card(row, "Total de Veículos",  "—", "🚛", VERDE)
        self.card_aptos    = self._card(row, "Aptos a Executar",   "—", "✅", VERDE)
        self.card_ignor    = self._card(row, "Ignorados",          "—", "⛔", AMARELO)
        self.card_semail   = self._card(row, "Sem E-mail",         "—", "⚠️", VERMELHO)
        for c in [self.card_total, self.card_aptos, self.card_ignor, self.card_semail]:
            c.pack(side="left", padx=8)

        ctk.CTkButton(
            page, text="🔄  Atualizar Dashboard",
            fg_color=VERDE, text_color=AZUL_ESCURO,
            font=ctk.CTkFont(size=13, weight="bold"), height=38,
            command=self._atualizar_dashboard
        ).pack(anchor="w", padx=30, pady=(0, 12))

        ctk.CTkFrame(page, height=1, fg_color="#21262d").pack(fill="x", padx=30, pady=8)

        ctk.CTkLabel(page, text="📋  Log do Sistema",
                     font=ctk.CTkFont(size=14, weight="bold"), text_color=BRANCO
                     ).pack(anchor="w", padx=30, pady=(0, 5))

        self.log_dash = self._make_log_box(page, height=260)
        self.log_dash.pack(fill="both", expand=True, padx=30, pady=(0, 20))

        self._log(self.log_dash, "Sistema iniciado com sucesso.")
        self._log(self.log_dash, "Clique em 'Atualizar Dashboard' para carregar os dados.")
        return page

    def _atualizar_dashboard(self):
        def run():
            try:
                arq = r"Z:\PV - Pos Vendas\EQUIPE\2026\Gestão\REVISÕES DE VEÍCULOS\REVISÕES.xlsx"
                if not os.path.exists(arq):
                    self._log(self.log_dash, f"❌ Arquivo não encontrado: {arq}")
                    return
                df = pd.read_excel(arq, sheet_name="PROSPEC FINAL")
                total   = len(df)
                aptos   = len(df[df["Status"] == 1])
                ignor   = len(df[df["Status"] != 1])
                semail  = len(df[df["E-mails"].isna() | (df["E-mails"].astype(str).str.strip() == "")])
                self._set_card_value(self.card_total,  str(total))
                self._set_card_value(self.card_aptos,  str(aptos))
                self._set_card_value(self.card_ignor,  str(ignor))
                self._set_card_value(self.card_semail, str(semail))
                self._log(self.log_dash,
                    f"✅ Dashboard atualizado | Total: {total} | Aptos: {aptos} | Ignorados: {ignor} | Sem e-mail: {semail}")
            except Exception as e:
                self._log(self.log_dash, f"❌ Erro: {e}")
        threading.Thread(target=run, daemon=True).start()

    # ============================================
    # KM ATUAL
    # ============================================
    def _build_km_atual(self):
        page = ctk.CTkFrame(self.main_area, fg_color="#0d1117", corner_radius=0)

        ctk.CTkLabel(page, text="📡  Atualização de KM Atual",
                     font=ctk.CTkFont(size=24, weight="bold"), text_color=BRANCO
                     ).pack(anchor="w", padx=30, pady=(25, 3))
        ctk.CTkLabel(page, text="Consulta a API Sascar e atualiza o KM atual dos veículos na planilha",
                     font=ctk.CTkFont(size=13), text_color="#888888"
                     ).pack(anchor="w", padx=30, pady=(0, 18))

        # Info
        info = ctk.CTkFrame(page, fg_color="#161b22", corner_radius=10)
        info.pack(fill="x", padx=30, pady=(0, 14))
        ctk.CTkLabel(info, text="ℹ️  O que este módulo faz:",
                     font=ctk.CTkFont(size=13, weight="bold"), text_color=BRANCO
                     ).pack(anchor="w", padx=15, pady=(10, 5))
        for txt in [
            "• Conecta à API Sascar com as credenciais configuradas",
            "• Consulta o KM atual de cada veículo pelo chassi",
            "• Atualiza a coluna 'Km Atual' na planilha REVISÕES.xlsx",
            "• Gera log de execução em Core/km_atual.log",
        ]:
            ctk.CTkLabel(info, text=txt, font=ctk.CTkFont(size=12), text_color="#aaaaaa"
                         ).pack(anchor="w", padx=25)
        ctk.CTkFrame(info, height=8, fg_color="transparent").pack()

        self.bar_km, self.lbl_km = self._make_progress(page)

        btn_row = ctk.CTkFrame(page, fg_color="transparent")
        btn_row.pack(anchor="w", padx=30, pady=(0, 14))
        ctk.CTkButton(btn_row, text="▶  Executar Atualização",
                      fg_color=VERDE, text_color=AZUL_ESCURO,
                      font=ctk.CTkFont(size=13, weight="bold"), height=40,
                      command=self._executar_km
                      ).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_row, text="🗑  Limpar Log",
                      fg_color="#21262d", text_color=BRANCO, height=40,
                      command=lambda: self._limpar_log(self.log_km)
                      ).pack(side="left")

        ctk.CTkLabel(page, text="📋  Log de Execução",
                     font=ctk.CTkFont(size=13, weight="bold"), text_color=BRANCO
                     ).pack(anchor="w", padx=30, pady=(0, 5))
        self.log_km = self._make_log_box(page, height=350)
        self.log_km.pack(fill="both", expand=True, padx=30, pady=(0, 20))
        return page

    def _executar_km(self):
        def run():
            self._log(self.log_km, "🔄 Iniciando atualização de KM Atual...")
            self.bar_km.set(0.1)
            self.lbl_km.configure(text="Importando módulo...")
            try:
                from Core.Km_atual import Km_atual
                old = sys.stdout; sys.stdout = io.StringIO()
                self.bar_km.set(0.4); self.lbl_km.configure(text="Conectando à API Sascar...")
                Km_atual()
                out = sys.stdout.getvalue(); sys.stdout = old
                for line in out.splitlines():
                    self._log(self.log_km, line)
                self.bar_km.set(1.0); self.lbl_km.configure(text="✅ Concluído!")
                self._log(self.log_km, "✅ KM Atual atualizado com sucesso!")
                self._log(self.log_dash, "✅ Módulo KM Atual executado.")
            except Exception as e:
                self.bar_km.set(0); self.lbl_km.configure(text="❌ Erro")
                self._log(self.log_km, f"❌ Erro: {e}")
        threading.Thread(target=run, daemon=True).start()

    # ============================================
    # PROSPECÇÃO
    # ============================================
    def _build_prospeccao(self):
        page = ctk.CTkFrame(self.main_area, fg_color="#0d1117", corner_radius=0)

        ctk.CTkLabel(page, text="📧  Prospecção de Clientes",
                     font=ctk.CTkFont(size=24, weight="bold"), text_color=BRANCO
                     ).pack(anchor="w", padx=30, pady=(25, 3))
        ctk.CTkLabel(page, text="Envia e-mails de agendamento de revisão para clientes com veículos aptos",
                     font=ctk.CTkFont(size=13), text_color="#888888"
                     ).pack(anchor="w", padx=30, pady=(0, 18))

        # Opções de envio
        opts = ctk.CTkFrame(page, fg_color="#161b22", corner_radius=10)
        opts.pack(fill="x", padx=30, pady=(0, 14))
        ctk.CTkLabel(opts, text="⚙️  Modo de Envio",
                     font=ctk.CTkFont(size=13, weight="bold"), text_color=BRANCO
                     ).pack(anchor="w", padx=15, pady=(10, 8))
        self.modo_envio = ctk.StringVar(value="N")
        rrow = ctk.CTkFrame(opts, fg_color="transparent")
        rrow.pack(anchor="w", padx=15, pady=(0, 12))
        ctk.CTkRadioButton(rrow, text="📝  Salvar como Rascunho",
                           variable=self.modo_envio, value="N",
                           text_color=BRANCO, fg_color=VERDE
                           ).pack(side="left", padx=(0, 25))
        ctk.CTkRadioButton(rrow, text="📤  Enviar Agora",
                           variable=self.modo_envio, value="S",
                           text_color=BRANCO, fg_color=VERDE
                           ).pack(side="left")

        self.bar_prosp, self.lbl_prosp = self._make_progress(page)

        btn_row = ctk.CTkFrame(page, fg_color="transparent")
        btn_row.pack(anchor="w", padx=30, pady=(0, 14))
        ctk.CTkButton(btn_row, text="▶  Executar Prospecção",
                      fg_color=VERDE, text_color=AZUL_ESCURO,
                      font=ctk.CTkFont(size=13, weight="bold"), height=40,
                      command=self._executar_prospeccao
                      ).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_row, text="👁  Pré-visualizar Clientes",
                      fg_color="#21262d", text_color=BRANCO, height=40,
                      command=self._preview_clientes
                      ).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_row, text="🗑  Limpar Log",
                      fg_color="#21262d", text_color=BRANCO, height=40,
                      command=lambda: self._limpar_log(self.log_prosp)
                      ).pack(side="left")

        ctk.CTkLabel(page, text="📋  Log de Execução",
                     font=ctk.CTkFont(size=13, weight="bold"), text_color=BRANCO
                     ).pack(anchor="w", padx=30, pady=(0, 5))
        self.log_prosp = self._make_log_box(page, height=330)
        self.log_prosp.pack(fill="both", expand=True, padx=30, pady=(0, 20))
        return page

    def _preview_clientes(self):
        try:
            arq = r"Z:\PV - Pos Vendas\EQUIPE\2026\Gestão\REVISÕES DE VEÍCULOS\REVISÕES.xlsx"
            df = pd.read_excel(arq, sheet_name="PROSPEC FINAL")
            aptos = df[df["Status"] == 1]
            self._log(self.log_prosp, f"👁 {len(aptos)} clientes aptos encontrados.")

            pop = ctk.CTkToplevel(self)
            pop.title("Pré-visualização — Clientes Aptos")
            pop.geometry("950x500")
            pop.configure(fg_color="#0d1117")
            pop.grab_set()

            ctk.CTkLabel(pop, text=f"👁  Clientes Aptos ({len(aptos)})",
                         font=ctk.CTkFont(size=16, weight="bold"), text_color=BRANCO
                         ).pack(pady=15)

            frame_t = ctk.CTkFrame(pop, fg_color="#161b22", corner_radius=10)
            frame_t.pack(fill="both", expand=True, padx=20, pady=(0, 20))

            style = ttk.Style()
            style.theme_use("default")
            style.configure("Treeview", background="#161b22", foreground="white",
                            fieldbackground="#161b22", rowheight=28)
            style.configure("Treeview.Heading", background="#041E42", foreground="white",
                            font=("Segoe UI", 11, "bold"))
            style.map("Treeview", background=[("selected", VERDE)],
                      foreground=[("selected", AZUL_ESCURO)])

            cols = ["Chassi", "Razão Social Cliente", "Placas", "Km Atual.1", "Km Revisão", "Revisão"]
            tree = ttk.Treeview(frame_t, columns=cols, show="headings", height=15)
            for col in cols:
                tree.heading(col, text=col)
                tree.column(col, width=145, anchor="center")
            sb = ttk.Scrollbar(frame_t, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=sb.set)
            for _, row in aptos.iterrows():
                tree.insert("", "end", values=[row.get(c, "") for c in cols])
            tree.pack(side="left", fill="both", expand=True, padx=10, pady=10)
            sb.pack(side="right", fill="y", pady=10)

        except Exception as e:
            self._log(self.log_prosp, f"❌ Erro ao pré-visualizar: {e}")

    def _executar_prospeccao(self):
        modo = self.modo_envio.get()
        def run():
            self._log(self.log_prosp,
                      f"🔄 Iniciando prospecção | Modo: {'📤 Enviar' if modo == 'S' else '📝 Rascunho'}...")
            self.bar_prosp.set(0.2); self.lbl_prosp.configure(text="Carregando planilha...")
            try:
                import builtins
                _in = builtins.input
                builtins.input = lambda _="": modo
                from Core.Prospeccao import Prospeccao
                old = sys.stdout; sys.stdout = io.StringIO()
                self.bar_prosp.set(0.5); self.lbl_prosp.configure(text="Enviando e-mails...")
                Prospeccao()
                out = sys.stdout.getvalue(); sys.stdout = old; builtins.input = _in
                for line in out.splitlines():
                    self._log(self.log_prosp, line)
                self.bar_prosp.set(1.0); self.lbl_prosp.configure(text="✅ Concluído!")
                self._log(self.log_prosp, "✅ Prospecção concluída!")
                self._log(self.log_dash, "✅ Módulo Prospecção executado.")
            except Exception as e:
                self.bar_prosp.set(0); self.lbl_prosp.configure(text="❌ Erro")
                self._log(self.log_prosp, f"❌ Erro: {e}")
        threading.Thread(target=run, daemon=True).start()

    # ============================================
    # HISTÓRICO
    # ============================================
    def _build_historico(self):
        page = ctk.CTkFrame(self.main_area, fg_color="#0d1117", corner_radius=0)

        ctk.CTkLabel(page, text="📋  Histórico de Seleção",
                     font=ctk.CTkFont(size=24, weight="bold"), text_color=BRANCO
                     ).pack(anchor="w", padx=30, pady=(25, 3))
        ctk.CTkLabel(page, text="Processa o arquivo SLK e atualiza o contador de seleções por chassi",
                     font=ctk.CTkFont(size=13), text_color="#888888"
                     ).pack(anchor="w", padx=30, pady=(0, 18))

        opts = ctk.CTkFrame(page, fg_color="#161b22", corner_radius=10)
        opts.pack(fill="x", padx=30, pady=(0, 14))
        ctk.CTkLabel(opts, text="⚙️  Atualização de Contador",
                     font=ctk.CTkFont(size=13, weight="bold"), text_color=BRANCO
                     ).pack(anchor="w", padx=15, pady=(10, 8))
        self.atualizar_cont = ctk.StringVar(value="N")
        rrow = ctk.CTkFrame(opts, fg_color="transparent")
        rrow.pack(anchor="w", padx=15, pady=(0, 12))
        ctk.CTkRadioButton(rrow, text="❌  Não atualizar contador",
                           variable=self.atualizar_cont, value="N",
                           text_color=BRANCO, fg_color=VERDE
                           ).pack(side="left", padx=(0, 25))
        ctk.CTkRadioButton(rrow, text="✅  Atualizar contador geral",
                           variable=self.atualizar_cont, value="S",
                           text_color=BRANCO, fg_color=VERDE
                           ).pack(side="left")

        self.bar_hist, self.lbl_hist = self._make_progress(page)

        btn_row = ctk.CTkFrame(page, fg_color="transparent")
        btn_row.pack(anchor="w", padx=30, pady=(0, 14))
        ctk.CTkButton(btn_row, text="▶  Processar Histórico",
                      fg_color=VERDE, text_color=AZUL_ESCURO,
                      font=ctk.CTkFont(size=13, weight="bold"), height=40,
                      command=self._executar_historico
                      ).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_row, text="👁  Ver Histórico",
                      fg_color="#21262d", text_color=BRANCO, height=40,
                      command=self._ver_historico
                      ).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_row, text="🗑  Limpar Log",
                      fg_color="#21262d", text_color=BRANCO, height=40,
                      command=lambda: self._limpar_log(self.log_hist)
                      ).pack(side="left")

        ctk.CTkLabel(page, text="📋  Log de Execução",
                     font=ctk.CTkFont(size=13, weight="bold"), text_color=BRANCO
                     ).pack(anchor="w", padx=30, pady=(0, 5))
        self.log_hist = self._make_log_box(page, height=330)
        self.log_hist.pack(fill="both", expand=True, padx=30, pady=(0, 20))
        return page

    def _ver_historico(self):
        try:
            arq = r"Z:\PV - Pos Vendas\EQUIPE\2026\Gestão\REVISÕES DE VEÍCULOS\HISTÓRICO DE VEÍCULOS\HISTORICO_SELECAO.xlsx"
            if not os.path.exists(arq):
                self._log(self.log_hist, "❌ Arquivo de histórico não encontrado.")
                return
            df = pd.read_excel(arq, sheet_name="PROSPECÇÃO")

            pop = ctk.CTkToplevel(self)
            pop.title("Histórico de Seleção de Veículos")
            pop.geometry("720x520")
            pop.configure(fg_color="#0d1117")
            pop.grab_set()

            ctk.CTkLabel(pop, text=f"📋  Histórico ({len(df)} registros)",
                         font=ctk.CTkFont(size=16, weight="bold"), text_color=BRANCO
                         ).pack(pady=15)

            # Busca
            srow = ctk.CTkFrame(pop, fg_color="transparent")
            srow.pack(fill="x", padx=20, pady=(0, 8))
            ctk.CTkLabel(srow, text="🔍 Buscar Chassi:", text_color=BRANCO).pack(side="left", padx=(0, 8))
            self._search_hist = ctk.StringVar()
            ctk.CTkEntry(srow, textvariable=self._search_hist, width=260).pack(side="left")

            frame_t = ctk.CTkFrame(pop, fg_color="#161b22", corner_radius=10)
            frame_t.pack(fill="both", expand=True, padx=20, pady=(0, 20))

            style = ttk.Style()
            style.theme_use("default")
            style.configure("Treeview", background="#161b22", foreground="white",
                            fieldbackground="#161b22", rowheight=28)
            style.configure("Treeview.Heading", background="#041E42", foreground="white",
                            font=("Segoe UI", 11, "bold"))
            style.map("Treeview", background=[("selected", VERDE)],
                      foreground=[("selected", AZUL_ESCURO)])

            tree = ttk.Treeview(frame_t, columns=["Chassi", "Seleções"], show="headings", height=16)
            tree.heading("Chassi", text="Chassi")
            tree.heading("Seleções", text="Nº de Seleções")
            tree.column("Chassi",   width=380, anchor="center")
            tree.column("Seleções", width=200, anchor="center")
            sb = ttk.Scrollbar(frame_t, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=sb.set)

            def populate(f=""):
                for i in tree.get_children(): tree.delete(i)
                for _, row in df.iterrows():
                    if f.lower() in str(row["Chassi"]).lower():
                        tree.insert("", "end", values=[row["Chassi"], int(row["Seleção"])])

            self._search_hist.trace("w", lambda *a: populate(self._search_hist.get()))
            populate()
            tree.pack(side="left", fill="both", expand=True, padx=10, pady=10)
            sb.pack(side="right", fill="y", pady=10)

        except Exception as e:
            self._log(self.log_hist, f"❌ Erro ao abrir histórico: {e}")

    def _executar_historico(self):
        escolha = self.atualizar_cont.get()
        def run():
            self._log(self.log_hist,
                      f"🔄 Processando histórico | Atualizar contador: {'Sim' if escolha == 'S' else 'Não'}...")
            self.bar_hist.set(0.2); self.lbl_hist.configure(text="Processando arquivo SLK...")
            try:
                import builtins
                _in = builtins.input
                builtins.input = lambda _="": escolha
                from Core.historico import historico
                old = sys.stdout; sys.stdout = io.StringIO()
                self.bar_hist.set(0.6); self.lbl_hist.configure(text="Atualizando histórico...")
                historico()
                out = sys.stdout.getvalue(); sys.stdout = old; builtins.input = _in
                for line in out.splitlines():
                    self._log(self.log_hist, line)
                self.bar_hist.set(1.0); self.lbl_hist.configure(text="✅ Concluído!")
                self._log(self.log_hist, "✅ Histórico processado com sucesso!")
                self._log(self.log_dash, "✅ Módulo Histórico executado.")
            except Exception as e:
                self.bar_hist.set(0); self.lbl_hist.configure(text="❌ Erro")
                self._log(self.log_hist, f"❌ Erro: {e}")
        threading.Thread(target=run, daemon=True).start()


# ============================================
# ENTRY POINT
# ============================================
if __name__ == "__main__":
    app = App()
    app.mainloop()
