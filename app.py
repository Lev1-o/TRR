import customtkinter as ctk
import threading
import sys
import io
import runpy
import builtins
from datetime import datetime
import pandas as pd

# ─── Tema ────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

PLANILHA = r"Z:\Revisões Preventivas\REVISÕES.xlsx"
ABA       = "PROSPEC FINAL"

# ─── Helpers ─────────────────────────────────────────────────────────────────
def timestamp():
    return datetime.now().strftime("[%H:%M:%S]")

def carregar_stats():
    try:
        df = pd.read_excel(PLANILHA, sheet_name=ABA)
        total    = len(df)
        aptos    = int((df["Status"] == 1).sum())
        ignorados= total - aptos
        sem_email= int(df["E-mail"].isna().sum()) if "E-mail" in df.columns else 0
        return total, aptos, ignorados, sem_email
    except Exception:
        return 0, 0, 0, 0

# ─── App principal ────────────────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("EQUIPO | Sistema de Revisões")
        self.geometry("1100x680")
        self.resizable(False, False)

        # Sidebar
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        ctk.CTkLabel(self.sidebar, text="EQUIPO", font=("Arial", 20, "bold"),
                     text_color="#1f8ef1").pack(pady=(30, 5))
        ctk.CTkLabel(self.sidebar, text="Sistema de Revisões",
                     font=("Arial", 10), text_color="gray").pack(pady=(0, 30))

        self.btn_dash    = self._sidebar_btn("🏠  Dashboard",   self._show_dashboard)
        self.btn_km      = self._sidebar_btn("📡  KM Atual",    self._show_km)
        self.btn_prosp   = self._sidebar_btn("📧  Prospecção",  self._show_prospeccao)
        self.btn_hist    = self._sidebar_btn("📋  Histórico",   self._show_historico)

        ctk.CTkLabel(self.sidebar, text=f"v1.0 – {datetime.now():%d/%m/%Y}",
                     font=("Arial", 9), text_color="gray").pack(side="bottom", pady=10)

        # Container central
        self.container = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.container.pack(side="left", fill="both", expand=True)

        self._build_dashboard()
        self._build_km()
        self._build_prospeccao()
        self._build_historico()

        self._show_dashboard()

    # ── Sidebar button ─────────────────────────────────────────────────────
    def _sidebar_btn(self, text, cmd):
        b = ctk.CTkButton(self.sidebar, text=text, anchor="w",
                          fg_color="transparent", hover_color="#2a2d3e",
                          font=("Arial", 13), command=cmd)
        b.pack(fill="x", padx=10, pady=4)
        return b

    def _hide_all(self):
        for w in [self.frm_dash, self.frm_km, self.frm_prosp, self.frm_hist]:
            w.pack_forget()

    # ── Log helper ─────────────────────────────────────────────────────────
    def _log(self, widget, msg):
        def _insert():
            widget.configure(state="normal")
            widget.insert("end", f"{timestamp()} {msg}\n")
            widget.see("end")
            widget.configure(state="disabled")
        self.after(0, _insert)

    # ══════════════════════════════════════════════════════════════════════
    # DASHBOARD
    # ══════════════════════════════════════════════════════════════════════
    def _build_dashboard(self):
        self.frm_dash = ctk.CTkFrame(self.container, corner_radius=0,
                                     fg_color="transparent")

        ctk.CTkLabel(self.frm_dash, text="Dashboard",
                     font=("Arial", 22, "bold")).pack(anchor="w", padx=30, pady=(25, 10))

        cards = ctk.CTkFrame(self.frm_dash, fg_color="transparent")
        cards.pack(fill="x", padx=30, pady=10)

        self.card_total    = self._card(cards, "🚛 Total de Veículos", "–")
        self.card_aptos    = self._card(cards, "✅ Aptos a Executar",  "–")
        self.card_ignorados= self._card(cards, "⛔ Ignorados",         "–")
        self.card_sememail = self._card(cards, "⚠️ Sem E-mail",        "–")

        ctk.CTkButton(self.frm_dash, text="🔄 Atualizar Dashboard",
                      command=self._atualizar_dashboard).pack(pady=5)

        ctk.CTkLabel(self.frm_dash, text="Log do sistema",
                     font=("Arial", 13, "bold")).pack(anchor="w", padx=30, pady=(20, 5))
        self.log_dash = ctk.CTkTextbox(self.frm_dash, height=250, state="disabled",
                                       font=("Consolas", 11))
        self.log_dash.pack(fill="x", padx=30, pady=(0, 20))

        self._atualizar_dashboard()

    def _card(self, parent, label, valor):
        f = ctk.CTkFrame(parent, width=210, height=90, corner_radius=12)
        f.pack(side="left", padx=8, pady=4)
        f.pack_propagate(False)
        ctk.CTkLabel(f, text=label, font=("Arial", 11), text_color="gray").pack(pady=(12, 2))
        lbl = ctk.CTkLabel(f, text=valor, font=("Arial", 26, "bold"))
        lbl.pack()
        return lbl

    def _atualizar_dashboard(self):
        total, aptos, ignorados, sem_email = carregar_stats()
        self.card_total.configure(text=str(total))
        self.card_aptos.configure(text=str(aptos))
        self.card_ignorados.configure(text=str(ignorados))
        self.card_sememail.configure(text=str(sem_email))
        self._log(self.log_dash, f"Dashboard atualizado — {total} veículos, {aptos} aptos.")

    def _show_dashboard(self):
        self._hide_all()
        self.frm_dash.pack(fill="both", expand=True)

    # ══════════════════════════════════════════════════════════════════════
    # KM ATUAL
    # ══════════════════════════════════════════════════════════════════════
    def _build_km(self):
        self.frm_km = ctk.CTkFrame(self.container, corner_radius=0,
                                   fg_color="transparent")

        ctk.CTkLabel(self.frm_km, text="KM Atual",
                     font=("Arial", 22, "bold")).pack(anchor="w", padx=30, pady=(25, 5))
        ctk.CTkLabel(self.frm_km, text="Atualiza o KM de cada veículo via API Sascar.",
                     font=("Arial", 12), text_color="gray").pack(anchor="w", padx=30)

        self.btn_exec_km = ctk.CTkButton(self.frm_km, text="▶  Executar KM Atual",
                                         width=220, command=self._executar_km)
        self.btn_exec_km.pack(pady=20)

        self.bar_km = ctk.CTkProgressBar(self.frm_km, width=500)
        self.bar_km.set(0)
        self.bar_km.pack(pady=5)

        self.lbl_km = ctk.CTkLabel(self.frm_km, text="Aguardando execução...",
                                   font=("Arial", 11), text_color="gray")
        self.lbl_km.pack(pady=5)

        ctk.CTkLabel(self.frm_km, text="Log de execução",
                     font=("Arial", 13, "bold")).pack(anchor="w", padx=30, pady=(15, 5))
        self.log_km = ctk.CTkTextbox(self.frm_km, height=330, state="disabled",
                                     font=("Consolas", 11))
        self.log_km.pack(fill="x", padx=30, pady=(0, 20))

    def _executar_km(self):
        def run():
            self.btn_exec_km.configure(state="disabled")
            self.bar_km.set(0.1)
            self.lbl_km.configure(text="Iniciando...")
            self._log(self.log_km, "🔄 Iniciando atualização de KM Atual...")

            old_stdout = sys.stdout
            sys.stdout = io.StringIO()
            try:
                self.bar_km.set(0.3)
                self.lbl_km.configure(text="Conectando à API Sascar...")
                runpy.run_module('Core.Km_atual', run_name='__main__')

                output = sys.stdout.getvalue()
                sys.stdout = old_stdout
                for line in output.splitlines():
                    if line.strip():
                        self._log(self.log_km, line)

                self.bar_km.set(1.0)
                self.lbl_km.configure(text="✅ Concluído com sucesso!")
                self._log(self.log_km, "✅ KM Atual atualizado com sucesso!")
                self._log(self.log_dash, "✅ KM Atual executado com sucesso.")
                self._atualizar_dashboard()

            except Exception as e:
                sys.stdout = old_stdout
                self.bar_km.set(0)
                self.lbl_km.configure(text="❌ Erro na execução")
                self._log(self.log_km, f"❌ Erro: {e}")
            finally:
                self.btn_exec_km.configure(state="normal")

        threading.Thread(target=run, daemon=True).start()

    def _show_km(self):
        self._hide_all()
        self.frm_km.pack(fill="both", expand=True)

    # ══════════════════════════════════════════════════════════════════════
    # PROSPECÇÃO
    # ══════════════════════════════════════════════════════════════════════
    def _build_prospeccao(self):
        self.frm_prosp = ctk.CTkFrame(self.container, corner_radius=0,
                                      fg_color="transparent")

        ctk.CTkLabel(self.frm_prosp, text="Prospecção de E-mails",
                     font=("Arial", 22, "bold")).pack(anchor="w", padx=30, pady=(25, 5))
        ctk.CTkLabel(self.frm_prosp,
                     text="Envia e-mails para veículos com Status = 1 na planilha.",
                     font=("Arial", 12), text_color="gray").pack(anchor="w", padx=30)

        row = ctk.CTkFrame(self.frm_prosp, fg_color="transparent")
        row.pack(pady=15)

        ctk.CTkLabel(row, text="Modo:", font=("Arial", 13)).pack(side="left", padx=(0, 10))
        self.modo_var = ctk.StringVar(value="enviar")
        ctk.CTkRadioButton(row, text="📤 Enviar Agora", variable=self.modo_var,
                           value="enviar").pack(side="left", padx=10)
        ctk.CTkRadioButton(row, text="📝 Salvar como Rascunho", variable=self.modo_var,
                           value="rascunho").pack(side="left", padx=10)

        self.btn_exec_prosp = ctk.CTkButton(self.frm_prosp, text="▶  Executar Prospecção",
                                            width=220, command=self._executar_prospeccao)
        self.btn_exec_prosp.pack(pady=5)

        ctk.CTkLabel(self.frm_prosp, text="Log de execução",
                     font=("Arial", 13, "bold")).pack(anchor="w", padx=30, pady=(15, 5))
        self.log_prosp = ctk.CTkTextbox(self.frm_prosp, height=360, state="disabled",
                                        font=("Consolas", 11))
        self.log_prosp.pack(fill="x", padx=30, pady=(0, 20))

    def _executar_prospeccao(self):
        modo = self.modo_var.get()

        def run():
            self.btn_exec_prosp.configure(state="disabled")
            self._log(self.log_prosp, f"🔄 Iniciando Prospecção — Modo: {modo.upper()}...")

            old_stdout = sys.stdout
            old_input  = builtins.input
            sys.stdout    = io.StringIO()
            # Simula resposta do input() que o módulo possa fazer
            builtins.input = lambda _="": "1" if modo == "enviar" else "2"

            try:
                runpy.run_module('Core.Prospeccao', run_name='__main__')

                output = sys.stdout.getvalue()
                sys.stdout   = old_stdout
                builtins.input = old_input

                for line in output.splitlines():
                    if line.strip():
                        self._log(self.log_prosp, line)

                self._log(self.log_prosp, "✅ Prospecção concluída com sucesso!")
                self._log(self.log_dash,  "✅ Prospecção executada com sucesso.")
                self._atualizar_dashboard()

            except Exception as e:
                sys.stdout     = old_stdout
                builtins.input = old_input
                self._log(self.log_prosp, f"❌ Erro: {e}")
            finally:
                self.btn_exec_prosp.configure(state="normal")

        threading.Thread(target=run, daemon=True).start()

    def _show_prospeccao(self):
        self._hide_all()
        self.frm_prosp.pack(fill="both", expand=True)

    # ══════════════════════════════════════════════════════════════════════
    # HISTÓRICO
    # ══════════════════════════════════════════════════════════════════════
    def _build_historico(self):
        self.frm_hist = ctk.CTkFrame(self.container, corner_radius=0,
                                     fg_color="transparent")

        ctk.CTkLabel(self.frm_hist, text="Histórico de Seleções",
                     font=("Arial", 22, "bold")).pack(anchor="w", padx=30, pady=(25, 5))
        ctk.CTkLabel(self.frm_hist,
                     text="Registra e conta quantas vezes cada chassi foi selecionado.",
                     font=("Arial", 12), text_color="gray").pack(anchor="w", padx=30)

        row = ctk.CTkFrame(self.frm_hist, fg_color="transparent")
        row.pack(pady=15)

        self.atualizar_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(row, text="Atualizar contador ao executar",
                        variable=self.atualizar_var).pack(side="left", padx=10)

        self.btn_exec_hist = ctk.CTkButton(self.frm_hist, text="▶  Executar Histórico",
                                           width=220, command=self._executar_historico)
        self.btn_exec_hist.pack(pady=5)

        ctk.CTkLabel(self.frm_hist, text="Log de execução",
                     font=("Arial", 13, "bold")).pack(anchor="w", padx=30, pady=(15, 5))
        self.log_hist = ctk.CTkTextbox(self.frm_hist, height=380, state="disabled",
                                       font=("Consolas", 11))
        self.log_hist.pack(fill="x", padx=30, pady=(0, 20))

    def _executar_historico(self):
        atualizar = self.atualizar_var.get()

        def run():
            self.btn_exec_hist.configure(state="disabled")
            self._log(self.log_hist,
                      f"🔄 Iniciando Histórico — {'Atualizando' if atualizar else 'Apenas visualizando'}...")

            old_stdout = sys.stdout
            old_input  = builtins.input
            sys.stdout     = io.StringIO()
            builtins.input = lambda _="": "s" if atualizar else "n"

            try:
                runpy.run_module('Core.historico', run_name='__main__')

                output = sys.stdout.getvalue()
                sys.stdout     = old_stdout
                builtins.input = old_input

                for line in output.splitlines():
                    if line.strip():
                        self._log(self.log_hist, line)

                self._log(self.log_hist, "✅ Histórico processado com sucesso!")
                self._log(self.log_dash,  "✅ Histórico executado.")

            except Exception as e:
                sys.stdout     = old_stdout
                builtins.input = old_input
                self._log(self.log_hist, f"❌ Erro: {e}")
            finally:
                self.btn_exec_hist.configure(state="normal")

        threading.Thread(target=run, daemon=True).start()

    def _show_historico(self):
        self._hide_all()
        self.frm_hist.pack(fill="both", expand=True)


# ─── Entry point ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
