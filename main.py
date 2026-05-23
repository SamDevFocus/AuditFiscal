import sys
import os
import threading
import queue
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk

sys.path.insert(0, os.path.dirname(__file__))
from database.db import (init_db, get_empresas, insert_empresa, delete_empresa,
                          save_importacao, save_divergencia, add_log, get_historico)
from services.comparator import load_estoque, load_forlions, compare
from exports.excel_export import export_excel

# ─── Theme ───────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

BG_DARK  = "#0F1923"
BG_CARD  = "#1A2638"
BG_INPUT = "#0D1B2A"
ACC_BLUE = "#2E86DE"
ACC_TEAL = "#00C9A7"
ACC_WARN = "#F39C12"
ACC_ERR  = "#E74C3C"
ACC_OK   = "#27AE60"
ACC_PURP = "#9B59B6"
TXT_PRI  = "#F0F4F8"
TXT_SEC  = "#8899A6"
BORDER   = "#2A3A4A"


# ─── Utility ─────────────────────────────────────────────────────────────────
def fmt_money(v):
    try: return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except: return "R$ 0,00"

def fmt_pct(v):
    try: return f"{float(v):.1f}%"
    except: return "0.0%"


# ═══════════════════════════════════════════════════════════════════════════════
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        init_db()
        self.title("AuditFiscal Pro — Sistema de Auditoria e Conferência")
        self.geometry("1400x860")
        self.minsize(1100, 700)
        self.configure(fg_color=BG_DARK)

        self.arquivo_estoque = tk.StringVar()
        self.arquivo_forlions = tk.StringVar()
        self.results = {}
        self.all_rows = []
        self.log_queue = queue.Queue()
        self.importacao_id = None

        self._build_sidebar()
        self._build_main()
        self._show_page("dashboard")
        self._process_log_queue()

    # ── Sidebar ───────────────────────────────────────────────────────────────
    def _build_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=220, fg_color=BG_CARD, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # Logo
        logo_frame = ctk.CTkFrame(self.sidebar, fg_color="#0D1B2A", height=70, corner_radius=0)
        logo_frame.pack(fill="x")
        ctk.CTkLabel(logo_frame, text="⚖  AuditFiscal", font=ctk.CTkFont("Arial", 18, "bold"),
                     text_color=ACC_TEAL).pack(pady=20)

        sep = ctk.CTkFrame(self.sidebar, height=1, fg_color=BORDER)
        sep.pack(fill="x")

        nav_items = [
            ("🏠", "Dashboard", "dashboard"),
            ("📂", "Importar Arquivos", "import"),
            ("📊", "Resultados", "results"),
            ("⚠", "Divergências", "divergencias"),
            ("🏢", "Empresas", "empresas"),
            ("📜", "Histórico", "historico"),
        ]
        self.nav_buttons = {}
        for icon, label, page in nav_items:
            btn = ctk.CTkButton(
                self.sidebar, text=f"  {icon}  {label}",
                fg_color="transparent", hover_color="#253545",
                anchor="w", height=42, corner_radius=0,
                font=ctk.CTkFont("Arial", 13),
                text_color=TXT_PRI,
                command=lambda p=page: self._show_page(p)
            )
            btn.pack(fill="x", padx=0)
            self.nav_buttons[page] = btn

        # Version at bottom
        ctk.CTkLabel(self.sidebar, text="v1.0.0 · 2025", font=ctk.CTkFont("Arial", 10),
                     text_color=TXT_SEC).pack(side="bottom", pady=10)

    def _show_page(self, page):
        for p, btn in self.nav_buttons.items():
            btn.configure(fg_color=ACC_BLUE if p == page else "transparent",
                          text_color="white" if p == page else TXT_PRI)
        for frame in self.pages.values():
            frame.pack_forget()
        self.pages[page].pack(fill="both", expand=True)

    # ── Main area ─────────────────────────────────────────────────────────────
    def _build_main(self):
        self.main = ctk.CTkFrame(self, fg_color=BG_DARK, corner_radius=0)
        self.main.pack(side="right", fill="both", expand=True)

        self.pages = {}
        for page in ["dashboard", "import", "results", "divergencias", "empresas", "historico"]:
            f = ctk.CTkFrame(self.main, fg_color=BG_DARK, corner_radius=0)
            self.pages[page] = f

        self._build_dashboard()
        self._build_import()
        self._build_results()
        self._build_divergencias()
        self._build_empresas()
        self._build_historico()

    # ── Dashboard ─────────────────────────────────────────────────────────────
    def _build_dashboard(self):
        page = self.pages["dashboard"]

        header = ctk.CTkFrame(page, fg_color=BG_CARD, height=60, corner_radius=0)
        header.pack(fill="x")
        ctk.CTkLabel(header, text="📊  Dashboard de Auditoria",
                     font=ctk.CTkFont("Arial", 18, "bold"), text_color=TXT_PRI).pack(side="left", padx=20, pady=15)
        self.dash_date_lbl = ctk.CTkLabel(header, text=datetime.now().strftime("Hoje: %d/%m/%Y %H:%M"),
                                          font=ctk.CTkFont("Arial", 11), text_color=TXT_SEC)
        self.dash_date_lbl.pack(side="right", padx=20)

        scroll = ctk.CTkScrollableFrame(page, fg_color=BG_DARK)
        scroll.pack(fill="both", expand=True, padx=10, pady=10)

        # KPI cards row 1
        row1 = ctk.CTkFrame(scroll, fg_color="transparent")
        row1.pack(fill="x", pady=(0, 10))
        self.kpi_vars = {}
        kpis1 = [
            ("total_forlions", "Notas ForLions", "📋", ACC_BLUE, "0"),
            ("total_estoque", "Entradas Estoque", "📦", ACC_TEAL, "0"),
            ("total_conciliadas", "Conciliadas", "✓", ACC_OK, "0"),
            ("total_divergencias", "Divergências", "⚠", ACC_ERR, "0"),
        ]
        for key, label, icon, color, default in kpis1:
            card = ctk.CTkFrame(row1, fg_color=BG_CARD, corner_radius=12)
            card.pack(side="left", fill="both", expand=True, padx=5)
            ctk.CTkFrame(card, height=4, fg_color=color, corner_radius=2).pack(fill="x")
            inner = ctk.CTkFrame(card, fg_color="transparent")
            inner.pack(fill="both", expand=True, padx=15, pady=12)
            top = ctk.CTkFrame(inner, fg_color="transparent")
            top.pack(fill="x")
            ctk.CTkLabel(top, text=icon, font=ctk.CTkFont("Arial", 20), text_color=color).pack(side="left")
            ctk.CTkLabel(top, text=label, font=ctk.CTkFont("Arial", 10), text_color=TXT_SEC).pack(side="right")
            var = tk.StringVar(value=default)
            self.kpi_vars[key] = var
            ctk.CTkLabel(inner, textvariable=var, font=ctk.CTkFont("Arial", 28, "bold"),
                         text_color=TXT_PRI).pack(anchor="w", pady=(5, 0))

        # Row 2
        row2 = ctk.CTkFrame(scroll, fg_color="transparent")
        row2.pack(fill="x", pady=(0, 10))
        kpis2 = [
            ("valor_total", "Valor Total Notas", "💰", ACC_BLUE, "R$ 0,00"),
            ("valor_conciliado", "Valor Conciliado", "✅", ACC_OK, "R$ 0,00"),
            ("conformidade", "Conformidade", "📈", ACC_TEAL, "0.0%"),
            ("total_despesas", "Despesas", "$", ACC_PURP, "0"),
        ]
        for key, label, icon, color, default in kpis2:
            card = ctk.CTkFrame(row2, fg_color=BG_CARD, corner_radius=12)
            card.pack(side="left", fill="both", expand=True, padx=5)
            ctk.CTkFrame(card, height=4, fg_color=color, corner_radius=2).pack(fill="x")
            inner = ctk.CTkFrame(card, fg_color="transparent")
            inner.pack(fill="both", expand=True, padx=15, pady=12)
            top = ctk.CTkFrame(inner, fg_color="transparent")
            top.pack(fill="x")
            ctk.CTkLabel(top, text=icon, font=ctk.CTkFont("Arial", 20), text_color=color).pack(side="left")
            ctk.CTkLabel(top, text=label, font=ctk.CTkFont("Arial", 10), text_color=TXT_SEC).pack(side="right")
            var = tk.StringVar(value=default)
            self.kpi_vars[key] = var
            ctk.CTkLabel(inner, textvariable=var, font=ctk.CTkFont("Arial", 24, "bold"),
                         text_color=TXT_PRI).pack(anchor="w", pady=(5, 0))

        # Status indicator
        status_card = ctk.CTkFrame(scroll, fg_color=BG_CARD, corner_radius=12)
        status_card.pack(fill="x", padx=5, pady=5)
        ctk.CTkLabel(status_card, text="📌  Status do Sistema",
                     font=ctk.CTkFont("Arial", 12, "bold"), text_color=TXT_PRI).pack(anchor="w", padx=15, pady=(12, 5))
        self.status_lbl = ctk.CTkLabel(status_card,
                                       text="Aguardando importação de arquivos. Use o menu 'Importar Arquivos' para iniciar.",
                                       font=ctk.CTkFont("Arial", 11), text_color=TXT_SEC)
        self.status_lbl.pack(anchor="w", padx=15, pady=(0, 12))

    def _update_dashboard(self):
        c = self.results.get("conciliadas", [])
        d = self.results.get("divergencias", []) + self.results.get("sem_entrada", [])
        e = self.results.get("despesas", [])
        dv = self.results.get("devolucoes", [])
        total_f = len(c) + len(d) + len(e) + len(dv)
        val_total = sum(r.get("valor_forlions", 0) or 0 for rows in self.results.values() for r in rows)
        val_conc = sum(r.get("valor_forlions", 0) or 0 for r in c)
        conformidade = (len(c) / total_f * 100) if total_f else 0

        self.kpi_vars["total_forlions"].set(str(total_f))
        self.kpi_vars["total_estoque"].set(str(len(self.df_estoque) if hasattr(self, 'df_estoque') and self.df_estoque is not None else 0))
        self.kpi_vars["total_conciliadas"].set(str(len(c)))
        self.kpi_vars["total_divergencias"].set(str(len(d)))
        self.kpi_vars["valor_total"].set(fmt_money(val_total))
        self.kpi_vars["valor_conciliado"].set(fmt_money(val_conc))
        self.kpi_vars["conformidade"].set(fmt_pct(conformidade))
        self.kpi_vars["total_despesas"].set(str(len(e)))

    # ── Import page ───────────────────────────────────────────────────────────
    def _build_import(self):
        page = self.pages["import"]

        header = ctk.CTkFrame(page, fg_color=BG_CARD, height=60, corner_radius=0)
        header.pack(fill="x")
        ctk.CTkLabel(header, text="📂  Importar Arquivos para Auditoria",
                     font=ctk.CTkFont("Arial", 18, "bold"), text_color=TXT_PRI).pack(side="left", padx=20, pady=15)

        content = ctk.CTkScrollableFrame(page, fg_color=BG_DARK)
        content.pack(fill="both", expand=True, padx=15, pady=15)

        # File cards
        files_row = ctk.CTkFrame(content, fg_color="transparent")
        files_row.pack(fill="x", pady=(0, 15))

        def make_file_card(parent, title, color, var, side):
            card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=12)
            card.pack(side=side, fill="both", expand=True, padx=5)
            ctk.CTkFrame(card, height=4, fg_color=color, corner_radius=2).pack(fill="x")
            inner = ctk.CTkFrame(card, fg_color="transparent")
            inner.pack(fill="both", expand=True, padx=20, pady=15)
            ctk.CTkLabel(inner, text=title, font=ctk.CTkFont("Arial", 12, "bold"),
                         text_color=TXT_PRI).pack(anchor="w")
            ctk.CTkLabel(inner, textvariable=var, font=ctk.CTkFont("Arial", 10),
                         text_color=TXT_SEC, wraplength=350).pack(anchor="w", pady=(4, 10))
            return inner

        inner1 = make_file_card(files_row, "📦  Arquivo de Entrada de Estoque", ACC_TEAL, self.arquivo_estoque, "left")
        ctk.CTkButton(inner1, text="📁  Selecionar Arquivo .xlsx",
                      fg_color=ACC_TEAL, hover_color="#00A589",
                      command=self._select_estoque).pack(anchor="w")

        inner2 = make_file_card(files_row, "📋  Arquivo de Notas ForLions", ACC_BLUE, self.arquivo_forlions, "right")
        ctk.CTkButton(inner2, text="📁  Selecionar Arquivo .xlsx",
                      fg_color=ACC_BLUE, hover_color="#1A6FBE",
                      command=self._select_forlions).pack(anchor="w")

        # Progress
        prog_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12)
        prog_card.pack(fill="x", padx=5, pady=(0, 10))
        inner_prog = ctk.CTkFrame(prog_card, fg_color="transparent")
        inner_prog.pack(fill="x", padx=20, pady=15)
        prog_top = ctk.CTkFrame(inner_prog, fg_color="transparent")
        prog_top.pack(fill="x")
        ctk.CTkLabel(prog_top, text="Progresso do Processamento",
                     font=ctk.CTkFont("Arial", 12, "bold"), text_color=TXT_PRI).pack(side="left")
        self.prog_lbl = ctk.CTkLabel(prog_top, text="0%",
                                     font=ctk.CTkFont("Arial", 11), text_color=ACC_TEAL)
        self.prog_lbl.pack(side="right")
        self.progress_bar = ctk.CTkProgressBar(inner_prog, height=8, progress_color=ACC_TEAL)
        self.progress_bar.pack(fill="x", pady=(8, 0))
        self.progress_bar.set(0)
        self.status_proc_lbl = ctk.CTkLabel(inner_prog, text="Aguardando arquivos...",
                                            font=ctk.CTkFont("Arial", 10), text_color=TXT_SEC)
        self.status_proc_lbl.pack(anchor="w", pady=(5, 0))

        # Buttons row
        btn_row = ctk.CTkFrame(content, fg_color="transparent")
        btn_row.pack(fill="x", padx=5, pady=(0, 10))
        self.btn_process = ctk.CTkButton(
            btn_row, text="▶  INICIAR AUDITORIA", height=44,
            font=ctk.CTkFont("Arial", 13, "bold"),
            fg_color=ACC_TEAL, hover_color="#00A589",
            command=self._start_processing
        )
        self.btn_process.pack(side="left", padx=(0, 10))
        self.btn_export = ctk.CTkButton(
            btn_row, text="⬇  Exportar Excel", height=44,
            font=ctk.CTkFont("Arial", 13, "bold"),
            fg_color=ACC_BLUE, hover_color="#1A6FBE",
            command=self._export_excel, state="disabled"
        )
        self.btn_export.pack(side="left")

        # Log console
        log_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12)
        log_card.pack(fill="both", expand=True, padx=5)
        ctk.CTkFrame(log_card, height=4, fg_color="#2A3A4A", corner_radius=2).pack(fill="x")
        log_header = ctk.CTkFrame(log_card, fg_color="transparent")
        log_header.pack(fill="x", padx=15, pady=(10, 5))
        ctk.CTkLabel(log_header, text="🖥  Console de Log em Tempo Real",
                     font=ctk.CTkFont("Arial", 11, "bold"), text_color=TXT_PRI).pack(side="left")
        ctk.CTkButton(log_header, text="Limpar", width=70, height=24,
                      font=ctk.CTkFont("Arial", 10), fg_color="#2A3A4A",
                      command=self._clear_log).pack(side="right")
        self.log_text = tk.Text(log_card, bg="#080F18", fg="#00FF88", insertbackground="white",
                                font=("Courier New", 9), relief="flat", bd=0, height=14,
                                selectbackground="#2A3A4A", wrap="word")
        log_scroll = ttk.Scrollbar(log_card, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log_text.tag_configure("OK",    foreground="#00FF88")
        self.log_text.tag_configure("INFO",  foreground="#7ECFFF")
        self.log_text.tag_configure("AVISO", foreground="#F39C12")
        self.log_text.tag_configure("ALERTA",foreground="#FF6B6B")
        self.log_text.tag_configure("ERRO",  foreground="#FF2D55")

    def _select_estoque(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if path:
            self.arquivo_estoque.set(path)
            self._log("INFO", f"Arquivo estoque selecionado: {os.path.basename(path)}")

    def _select_forlions(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if path:
            self.arquivo_forlions.set(path)
            self._log("INFO", f"Arquivo ForLions selecionado: {os.path.basename(path)}")

    def _log(self, nivel, msg):
        self.log_queue.put((nivel, msg))

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _process_log_queue(self):
        try:
            while True:
                nivel, msg = self.log_queue.get_nowait()
                ts = datetime.now().strftime("%H:%M:%S")
                self.log_text.configure(state="normal")
                self.log_text.insert("end", f"[{ts}] [{nivel}] {msg}\n", nivel)
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._process_log_queue)

    def _start_processing(self):
        fe = self.arquivo_estoque.get()
        ff = self.arquivo_forlions.get()
        if not fe or not os.path.exists(fe):
            messagebox.showerror("Erro", "Selecione o arquivo de Entrada de Estoque.")
            return
        if not ff or not os.path.exists(ff):
            messagebox.showerror("Erro", "Selecione o arquivo de Notas ForLions.")
            return

        self.btn_process.configure(state="disabled", text="⏳  Processando...")
        self.progress_bar.set(0)
        self.prog_lbl.configure(text="0%")
        self._log("INFO", "═" * 60)
        self._log("INFO", "INICIANDO AUDITORIA FISCAL")
        self._log("INFO", "═" * 60)
        threading.Thread(target=self._process_thread, args=(fe, ff), daemon=True).start()

    def _process_thread(self, fe, ff):
        try:
            self.df_estoque, err = load_estoque(fe, self._log)
            if err:
                self._log("ERRO", f"Falha ao ler estoque: {err}")
                self.after(0, lambda: self.btn_process.configure(state="normal", text="▶  INICIAR AUDITORIA"))
                return

            self.df_forlions, err = load_forlions(ff, self._log)
            if err:
                self._log("ERRO", f"Falha ao ler ForLions: {err}")
                self.after(0, lambda: self.btn_process.configure(state="normal", text="▶  INICIAR AUDITORIA"))
                return

            total = len(self.df_forlions)

            def progress_cb(current, total):
                pct = current / total if total else 0
                self.after(0, lambda: self.progress_bar.set(pct))
                self.after(0, lambda: self.prog_lbl.configure(text=f"{pct*100:.0f}%"))
                self.after(0, lambda: self.status_proc_lbl.configure(text=f"Processando nota {current} de {total}..."))

            self.results = compare(self.df_estoque, self.df_forlions, self._log, progress_cb)

            # Save to DB
            c  = len(self.results.get("conciliadas", []))
            d  = len(self.results.get("divergencias", [])) + len(self.results.get("sem_entrada", []))
            e  = len(self.results.get("despesas", []))
            dv = len(self.results.get("devolucoes", []))
            val_total = sum(r.get("valor_forlions", 0) or 0 for rows in self.results.values() for r in rows)
            val_conc  = sum(r.get("valor_forlions", 0) or 0 for r in self.results.get("conciliadas", []))
            self.importacao_id = save_importacao({
                "arquivo_estoque": os.path.basename(fe),
                "arquivo_forlions": os.path.basename(ff),
                "total_estoque": len(self.df_estoque),
                "total_forlions": total,
                "total_conciliadas": c,
                "total_divergencias": d,
                "total_despesas": e,
                "total_devolucoes": dv,
                "valor_total_notas": val_total,
                "valor_conciliado": val_conc,
            })
            for item in self.results.get("divergencias", []) + self.results.get("sem_entrada", []):
                save_divergencia({**item, "importacao_id": self.importacao_id})

            self._log("OK", "═" * 60)
            self._log("OK", "AUDITORIA CONCLUÍDA COM SUCESSO!")
            self._log("OK", f"  Conciliadas: {c}  |  Divergências: {d}  |  Despesas: {e}")
            self._log("OK", "═" * 60)

            self.after(0, self._on_process_done)
        except Exception as ex:
            self._log("ERRO", f"Erro inesperado: {ex}")
            import traceback
            self._log("ERRO", traceback.format_exc())
            self.after(0, lambda: self.btn_process.configure(state="normal", text="▶  INICIAR AUDITORIA"))

    def _on_process_done(self):
        self.btn_process.configure(state="normal", text="▶  INICIAR AUDITORIA")
        self.btn_export.configure(state="normal")
        self.progress_bar.set(1)
        self.prog_lbl.configure(text="100%")
        self.status_proc_lbl.configure(text="✓ Processamento concluído!")
        self._update_dashboard()
        self._populate_results()
        self._populate_divergencias()
        messagebox.showinfo("Concluído", "Auditoria finalizada! Verifique os resultados no menu 'Resultados'.")

    def _export_excel(self):
        if not self.results:
            messagebox.showwarning("Aviso", "Nenhum resultado para exportar.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"auditoria_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
        if not path:
            return
        c  = len(self.results.get("conciliadas", []))
        d  = len(self.results.get("divergencias", [])) + len(self.results.get("sem_entrada", []))
        e  = len(self.results.get("despesas", []))
        val_total = sum(r.get("valor_forlions", 0) or 0 for rows in self.results.values() for r in rows)
        val_conc  = sum(r.get("valor_forlions", 0) or 0 for r in self.results.get("conciliadas", []))
        conf = (c / (c + d + e) * 100) if (c + d + e) else 0
        export_excel(self.results, path, {
            "total_forlions": c + d + e,
            "total_estoque": len(self.df_estoque) if hasattr(self, 'df_estoque') and self.df_estoque is not None else 0,
            "total_conciliadas": c,
            "total_divergencias": d,
            "total_despesas": e,
            "total_devolucoes": len(self.results.get("devolucoes", [])),
            "valor_total_notas": val_total,
            "valor_conciliado": val_conc,
            "conformidade": conf,
        })
        messagebox.showinfo("Sucesso", f"Relatório exportado:\n{path}")

    # ── Results page ──────────────────────────────────────────────────────────
    def _build_results(self):
        page = self.pages["results"]

        header = ctk.CTkFrame(page, fg_color=BG_CARD, height=60, corner_radius=0)
        header.pack(fill="x")
        ctk.CTkLabel(header, text="📊  Resultados da Auditoria",
                     font=ctk.CTkFont("Arial", 18, "bold"), text_color=TXT_PRI).pack(side="left", padx=20, pady=15)

        controls = ctk.CTkFrame(page, fg_color=BG_CARD, height=50, corner_radius=0)
        controls.pack(fill="x")
        ctk.CTkLabel(controls, text="Filtrar:", text_color=TXT_SEC,
                     font=ctk.CTkFont("Arial", 11)).pack(side="left", padx=(15, 5), pady=12)
        self.result_filter = ctk.CTkComboBox(controls, values=["Todos", "Conciliadas", "Divergência de Valor",
                                                                  "Sem Entrada", "Despesa", "Devolução"],
                                              width=180, command=self._filter_results)
        self.result_filter.pack(side="left", pady=10, padx=5)
        self.result_search = ctk.CTkEntry(controls, placeholder_text="🔍  Buscar nota/fornecedor...",
                                          width=260)
        self.result_search.pack(side="left", pady=10, padx=10)
        self.result_search.bind("<KeyRelease>", lambda e: self._filter_results())
        ctk.CTkLabel(controls, textvariable=tk.StringVar(value=""), text_color=TXT_SEC).pack(side="left")
        self.result_count_lbl = ctk.CTkLabel(controls, text="0 registros",
                                              font=ctk.CTkFont("Arial", 10), text_color=TXT_SEC)
        self.result_count_lbl.pack(side="right", padx=15)

        tree_frame = ctk.CTkFrame(page, fg_color=BG_CARD)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Audit.Treeview",
                         background=BG_INPUT, foreground=TXT_PRI,
                         rowheight=28, fieldbackground=BG_INPUT,
                         font=("Arial", 10), borderwidth=0)
        style.configure("Audit.Treeview.Heading",
                         background=BG_CARD, foreground=TXT_PRI,
                         font=("Arial", 10, "bold"), relief="flat")
        style.map("Audit.Treeview", background=[("selected", "#253545")])

        cols = ("status", "numero", "razao_social", "data", "valor_forlions", "valor_estoque", "classificacao", "risco")
        self.result_tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                         style="Audit.Treeview")
        headers = {"status": ("Status", 120), "numero": ("Nº Nota", 90), "razao_social": ("Razão Social", 260),
                   "data": ("Data Emissão", 100), "valor_forlions": ("Valor ForLions", 110),
                   "valor_estoque": ("Valor Estoque", 110), "classificacao": ("Classificação", 120),
                   "risco": ("Risco", 70)}
        for col, (heading, width) in headers.items():
            self.result_tree.heading(col, text=heading)
            self.result_tree.column(col, width=width, anchor="center" if col not in ("razao_social",) else "w")

        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.result_tree.pack(fill="both", expand=True)

        self.result_tree.tag_configure("Conciliada", foreground="#27AE60")
        self.result_tree.tag_configure("Divergência de Valor", foreground="#E67E22")
        self.result_tree.tag_configure("Sem Entrada", foreground="#E74C3C")
        self.result_tree.tag_configure("Despesa", foreground="#3498DB")
        self.result_tree.tag_configure("Devolução", foreground="#9B59B6")

    def _populate_results(self):
        self.all_rows = []
        for category in ["conciliadas", "divergencias", "sem_entrada", "despesas", "devolucoes"]:
            for item in self.results.get(category, []):
                self.all_rows.append(item)
        self._filter_results()

    def _filter_results(self, *args):
        filt = self.result_filter.get() if hasattr(self, 'result_filter') else "Todos"
        search = self.result_search.get().upper().strip() if hasattr(self, 'result_search') else ""
        self.result_tree.delete(*self.result_tree.get_children())
        count = 0
        for item in self.all_rows:
            status = item.get("status", "")
            if filt != "Todos" and filt not in status:
                continue
            rs = (item.get("razao_social", "") or "").upper()
            num = (item.get("numero_nota", "") or "").upper()
            if search and search not in rs and search not in num:
                continue
            vf = fmt_money(item.get("valor_forlions", 0))
            ve = fmt_money(item.get("valor_estoque", 0)) if item.get("valor_estoque") is not None else "-"
            tag = status.replace(" ", "_") if status in ["Conciliada", "Sem Entrada", "Despesa", "Devolução"] else status
            self.result_tree.insert("", "end", values=(
                status, item.get("numero_nota", ""), item.get("razao_social", ""),
                item.get("data_emissao", ""), vf, ve, item.get("classificacao", ""), item.get("risco", "")
            ), tags=(status,))
            count += 1
        self.result_count_lbl.configure(text=f"{count} registros")

    # ── Divergências page ─────────────────────────────────────────────────────
    def _build_divergencias(self):
        page = self.pages["divergencias"]

        header = ctk.CTkFrame(page, fg_color=BG_CARD, height=60, corner_radius=0)
        header.pack(fill="x")
        ctk.CTkLabel(header, text="⚠  Divergências — Classificação e Revisão",
                     font=ctk.CTkFont("Arial", 18, "bold"), text_color=TXT_PRI).pack(side="left", padx=20, pady=15)

        paned = ctk.CTkFrame(page, fg_color=BG_DARK)
        paned.pack(fill="both", expand=True, padx=10, pady=10)

        # Table
        left = ctk.CTkFrame(paned, fg_color=BG_CARD, corner_radius=12)
        left.pack(side="left", fill="both", expand=True)

        cols2 = ("numero", "razao_social", "valor_forlions", "status", "classificacao", "risco")
        self.div_tree = ttk.Treeview(left, columns=cols2, show="headings", style="Audit.Treeview")
        hdrs2 = {"numero": ("Nº Nota", 90), "razao_social": ("Razão Social", 220),
                 "valor_forlions": ("Valor", 100), "status": ("Status", 120),
                 "classificacao": ("Classificação", 110), "risco": ("Risco", 70)}
        for col, (heading, width) in hdrs2.items():
            self.div_tree.heading(col, text=heading)
            self.div_tree.column(col, width=width, anchor="center" if col != "razao_social" else "w")
        sb2 = ttk.Scrollbar(left, orient="vertical", command=self.div_tree.yview)
        self.div_tree.configure(yscrollcommand=sb2.set)
        sb2.pack(side="right", fill="y")
        self.div_tree.pack(fill="both", expand=True, padx=5, pady=5)
        self.div_tree.bind("<<TreeviewSelect>>", self._on_div_select)

        # Detail panel
        right = ctk.CTkFrame(paned, fg_color=BG_CARD, corner_radius=12, width=300)
        right.pack(side="right", fill="y", padx=(10, 0))
        right.pack_propagate(False)

        ctk.CTkLabel(right, text="Classificar Divergência",
                     font=ctk.CTkFont("Arial", 12, "bold"), text_color=TXT_PRI).pack(padx=15, pady=(15, 5), anchor="w")

        self.div_detail_num = ctk.CTkLabel(right, text="Selecione uma nota",
                                            font=ctk.CTkFont("Arial", 10), text_color=TXT_SEC)
        self.div_detail_num.pack(padx=15, anchor="w")
        self.div_detail_val = ctk.CTkLabel(right, text="",
                                            font=ctk.CTkFont("Arial", 12, "bold"), text_color=ACC_ERR)
        self.div_detail_val.pack(padx=15, anchor="w", pady=(0, 10))

        ctk.CTkLabel(right, text="Classificação:", text_color=TXT_SEC, font=ctk.CTkFont("Arial", 10)).pack(padx=15, anchor="w")
        self.div_class_combo = ctk.CTkComboBox(right, values=["Pendente", "Erro", "Despesa", "Devolução",
                                                                "Serviço", "Nota Ignorada", "Outro"], width=260)
        self.div_class_combo.pack(padx=15, pady=(2, 10))

        ctk.CTkLabel(right, text="Status:", text_color=TXT_SEC, font=ctk.CTkFont("Arial", 10)).pack(padx=15, anchor="w")
        self.div_status_combo = ctk.CTkComboBox(right, values=["Pendente", "Em Revisão", "Aprovada", "Rejeitada"], width=260)
        self.div_status_combo.pack(padx=15, pady=(2, 10))

        ctk.CTkLabel(right, text="Observações:", text_color=TXT_SEC, font=ctk.CTkFont("Arial", 10)).pack(padx=15, anchor="w")
        self.div_obs_text = ctk.CTkTextbox(right, height=100, width=260, fg_color=BG_INPUT)
        self.div_obs_text.pack(padx=15, pady=(2, 10))

        ctk.CTkButton(right, text="💾  Salvar Classificação", fg_color=ACC_TEAL,
                      hover_color="#00A589", command=self._save_div_classification).pack(padx=15, fill="x", pady=(0, 5))
        ctk.CTkButton(right, text="🔄  Limpar Seleção", fg_color="#2A3A4A",
                      hover_color="#3A4A5A", command=self._clear_div_selection).pack(padx=15, fill="x")

        self._selected_div_iid = None

    def _populate_divergencias(self):
        self.div_tree.delete(*self.div_tree.get_children())
        all_div = self.results.get("divergencias", []) + self.results.get("sem_entrada", [])
        for item in all_div:
            self.div_tree.insert("", "end", values=(
                item.get("numero_nota", ""), item.get("razao_social", ""),
                fmt_money(item.get("valor_forlions", 0)), item.get("status", ""),
                item.get("classificacao", ""), item.get("risco", "")
            ), tags=(item.get("status", ""),))

    def _on_div_select(self, event):
        sel = self.div_tree.selection()
        if not sel:
            return
        self._selected_div_iid = sel[0]
        values = self.div_tree.item(sel[0])["values"]
        self.div_detail_num.configure(text=f"Nota: {values[0]}  |  {values[1]}")
        self.div_detail_val.configure(text=str(values[2]))
        self.div_class_combo.set(str(values[4]))
        self.div_status_combo.set(str(values[3]))

    def _save_div_classification(self):
        if not self._selected_div_iid:
            return
        cls = self.div_class_combo.get()
        status = self.div_status_combo.get()
        obs = self.div_obs_text.get("1.0", "end").strip()
        values = list(self.div_tree.item(self._selected_div_iid)["values"])
        values[3] = status
        values[4] = cls
        self.div_tree.item(self._selected_div_iid, values=values)
        self._log("INFO", f"Classificação salva: NF {values[0]} → {cls} ({status})")

    def _clear_div_selection(self):
        self._selected_div_iid = None
        self.div_detail_num.configure(text="Selecione uma nota")
        self.div_detail_val.configure(text="")

    # ── Empresas page ─────────────────────────────────────────────────────────
    def _build_empresas(self):
        page = self.pages["empresas"]

        header = ctk.CTkFrame(page, fg_color=BG_CARD, height=60, corner_radius=0)
        header.pack(fill="x")
        ctk.CTkLabel(header, text="🏢  Cadastro de Empresas e Fornecedores",
                     font=ctk.CTkFont("Arial", 18, "bold"), text_color=TXT_PRI).pack(side="left", padx=20, pady=15)

        content = ctk.CTkFrame(page, fg_color=BG_DARK)
        content.pack(fill="both", expand=True, padx=10, pady=10)

        # Form card
        form_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12)
        form_card.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(form_card, text="Novo Cadastro", font=ctk.CTkFont("Arial", 12, "bold"),
                     text_color=TXT_PRI).pack(anchor="w", padx=15, pady=(12, 5))
        form_inner = ctk.CTkFrame(form_card, fg_color="transparent")
        form_inner.pack(fill="x", padx=15, pady=(0, 12))

        fields_row1 = ctk.CTkFrame(form_inner, fg_color="transparent")
        fields_row1.pack(fill="x", pady=(0, 8))
        self.emp_nome_fantasia = self._form_entry(fields_row1, "Nome Fantasia", "left")
        self.emp_razao_social = self._form_entry(fields_row1, "Razão Social *", "left")
        self.emp_cnpj = self._form_entry(fields_row1, "CNPJ", "left")

        fields_row2 = ctk.CTkFrame(form_inner, fg_color="transparent")
        fields_row2.pack(fill="x")
        self.emp_tipo = ctk.CTkComboBox(fields_row2, values=["Fornecedor", "Despesa", "Devolução", "Serviço", "Outro"], width=160)
        self.emp_tipo.pack(side="left", padx=(0, 10))
        ctk.CTkButton(fields_row2, text="➕  Adicionar", fg_color=ACC_TEAL, hover_color="#00A589",
                      command=self._add_empresa).pack(side="left")
        ctk.CTkButton(fields_row2, text="🗑  Remover Selecionado", fg_color=ACC_ERR, hover_color="#C0392B",
                      command=self._del_empresa).pack(side="left", padx=10)
        ctk.CTkButton(fields_row2, text="🔄  Atualizar Lista", fg_color="#2A3A4A",
                      command=self._load_empresas).pack(side="right")

        # Table
        table_card = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=12)
        table_card.pack(fill="both", expand=True)
        cols3 = ("id", "nome_fantasia", "razao_social", "cnpj", "tipo_padrao")
        self.emp_tree = ttk.Treeview(table_card, columns=cols3, show="headings", style="Audit.Treeview")
        hdrs3 = {"id": ("ID", 40), "nome_fantasia": ("Nome Fantasia", 180),
                 "razao_social": ("Razão Social", 260), "cnpj": ("CNPJ", 140), "tipo_padrao": ("Tipo", 100)}
        for col, (heading, width) in hdrs3.items():
            self.emp_tree.heading(col, text=heading)
            self.emp_tree.column(col, width=width, anchor="w" if col in ("nome_fantasia", "razao_social") else "center")
        sb3 = ttk.Scrollbar(table_card, orient="vertical", command=self.emp_tree.yview)
        self.emp_tree.configure(yscrollcommand=sb3.set)
        sb3.pack(side="right", fill="y")
        self.emp_tree.pack(fill="both", expand=True, padx=5, pady=5)
        self._load_empresas()

    def _form_entry(self, parent, label, side):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(side=side, fill="x", expand=True, padx=(0, 10))
        ctk.CTkLabel(frame, text=label, font=ctk.CTkFont("Arial", 10), text_color=TXT_SEC).pack(anchor="w")
        entry = ctk.CTkEntry(frame, fg_color=BG_INPUT, border_color=BORDER)
        entry.pack(fill="x")
        return entry

    def _add_empresa(self):
        razao = self.emp_razao_social.get().strip()
        if not razao:
            messagebox.showwarning("Aviso", "Razão Social é obrigatória.")
            return
        result = insert_empresa(
            self.emp_nome_fantasia.get().strip(), razao,
            self.emp_cnpj.get().strip(), self.emp_tipo.get()
        )
        if result is True:
            self._load_empresas()
            for e in [self.emp_nome_fantasia, self.emp_razao_social, self.emp_cnpj]:
                e.delete(0, "end")
        else:
            messagebox.showerror("Erro", f"Falha ao salvar: {result}")

    def _del_empresa(self):
        sel = self.emp_tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione uma empresa para remover.")
            return
        emp_id = self.emp_tree.item(sel[0])["values"][0]
        if messagebox.askyesno("Confirmar", "Remover empresa selecionada?"):
            delete_empresa(emp_id)
            self._load_empresas()

    def _load_empresas(self):
        self.emp_tree.delete(*self.emp_tree.get_children())
        for emp in get_empresas():
            self.emp_tree.insert("", "end", values=(
                emp["id"], emp["nome_fantasia"] or "", emp["razao_social"],
                emp["cnpj"] or "", emp["tipo_padrao"]
            ))

    # ── Histórico page ────────────────────────────────────────────────────────
    def _build_historico(self):
        page = self.pages["historico"]

        header = ctk.CTkFrame(page, fg_color=BG_CARD, height=60, corner_radius=0)
        header.pack(fill="x")
        ctk.CTkLabel(header, text="📜  Histórico de Importações",
                     font=ctk.CTkFont("Arial", 18, "bold"), text_color=TXT_PRI).pack(side="left", padx=20, pady=15)
        ctk.CTkButton(header, text="🔄  Atualizar", width=100, fg_color=ACC_BLUE,
                      command=self._load_historico).pack(side="right", padx=15, pady=12)

        card = ctk.CTkFrame(page, fg_color=BG_CARD, corner_radius=12)
        card.pack(fill="both", expand=True, padx=10, pady=10)

        cols4 = ("id", "criado_em", "arquivo_estoque", "arquivo_forlions",
                 "total_forlions", "total_conciliadas", "total_divergencias", "valor_total")
        self.hist_tree = ttk.Treeview(card, columns=cols4, show="headings", style="Audit.Treeview")
        hdrs4 = {"id": ("ID", 40), "criado_em": ("Data/Hora", 130),
                 "arquivo_estoque": ("Arquivo Estoque", 200), "arquivo_forlions": ("Arquivo ForLions", 200),
                 "total_forlions": ("Total NF", 70), "total_conciliadas": ("Concil.", 70),
                 "total_divergencias": ("Diverg.", 70), "valor_total": ("Valor Total", 120)}
        for col, (heading, width) in hdrs4.items():
            self.hist_tree.heading(col, text=heading)
            self.hist_tree.column(col, width=width, anchor="center" if col not in ("arquivo_estoque","arquivo_forlions") else "w")
        sb4 = ttk.Scrollbar(card, orient="vertical", command=self.hist_tree.yview)
        self.hist_tree.configure(yscrollcommand=sb4.set)
        sb4.pack(side="right", fill="y")
        self.hist_tree.pack(fill="both", expand=True, padx=5, pady=5)
        self._load_historico()

    def _load_historico(self):
        self.hist_tree.delete(*self.hist_tree.get_children())
        for h in get_historico():
            self.hist_tree.insert("", "end", values=(
                h["id"], h["criado_em"][:16], h["arquivo_estoque"], h["arquivo_forlions"],
                h["total_forlions"], h["total_conciliadas"], h["total_divergencias"],
                fmt_money(h["valor_total_notas"])
            ))


# ─── Entry point ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
