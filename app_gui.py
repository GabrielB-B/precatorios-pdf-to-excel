from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import traceback
import unicodedata
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from leitor_pdf import run_extraction
from validar_planilhas import run_validation


APP_TITLE = "Extrator de Precatórios"
APP_ID = "gabriel.bispo.extrator.precatorios"

COLORS = {
    "ink": "#0f1a24",
    "ink_soft": "#162534",
    "ink_panel": "#1a2d40",
    "paper": "#f4efe5",
    "paper_soft": "#efe5d3",
    "paper_line": "#d9c9ac",
    "text_dark": "#172331",
    "text_soft": "#516274",
    "gold": "#c79734",
    "gold_soft": "#f1d697",
    "olive": "#68735f",
    "success": "#2c7a5a",
    "danger": "#a84b3f",
    "log_bg": "#12202d",
    "log_line": "#284055",
}

FONTS = {
    "display": ("Georgia", 24, "bold"),
    "display_small": ("Georgia", 18, "bold"),
    "title": ("Segoe UI Semibold", 11),
    "body": ("Segoe UI", 10),
    "body_small": ("Segoe UI", 9),
    "metric": ("Georgia", 16, "bold"),
    "button": ("Segoe UI Semibold", 10),
    "log": ("Consolas", 10),
}


def normalize_ascii(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", text)
    ascii_text = normalized.encode("ascii", "ignore").decode("ascii")
    return ascii_text.upper()


def project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_root() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(getattr(sys, "_MEIPASS"))
    return Path(__file__).resolve().parent


def resource_path(*parts: str) -> Path:
    return resource_root().joinpath(*parts)


def first_file(directory: Path | None, pattern: str) -> str:
    if directory is None or not directory.exists():
        return ""
    matches = sorted(path for path in directory.glob(pattern) if path.is_file() and not path.name.startswith("~$"))
    return str(matches[0].resolve()) if matches else ""


def find_validation_directory(root: Path) -> Path | None:
    for child in sorted(root.iterdir()):
        if child.is_dir() and normalize_ascii(child.name).startswith("VALIDA"):
            return child
    return None


def timestamp_suffix() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def open_path(target: Path) -> None:
    if sys.platform == "win32":
        os.startfile(str(target))
        return
    if sys.platform == "darwin":
        subprocess.run(["open", str(target)], check=False)
        return
    subprocess.run(["xdg-open", str(target)], check=False)


class ExtractorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()

        self.root_dir = project_root()
        self.validation_dir = find_validation_directory(self.root_dir)
        self.log_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker_thread: threading.Thread | None = None

        manual_default = first_file(self.validation_dir, "*.xlsx")
        self.pdf_var = tk.StringVar(value=first_file(self.root_dir / "entrada", "*.pdf"))
        self.manual_var = tk.StringVar(value=manual_default)
        self.output_var = tk.StringVar(value=str((self.root_dir / "saida").resolve()))
        self.validate_var = tk.BooleanVar(value=bool(manual_default))
        self.status_var = tk.StringVar(value="Pronto para consolidar um relatório.")
        self.metric_pdf_var = tk.StringVar(value="Aguardando PDF")
        self.metric_records_var = tk.StringVar(value="Sem leitura")
        self.metric_validation_var = tk.StringVar(value="Validação opcional")
        self.metric_quality_var = tk.StringVar(value="Sem execução")
        self.last_output_dir = Path(self.output_var.get())

        self._window_icon_photo: tk.PhotoImage | None = None
        self._brand_photo: tk.PhotoImage | None = None

        self._configure_window()
        self._configure_window_icon()
        self._build_layout()
        self.after(150, self._drain_queue)

    def _configure_window(self) -> None:
        self.title(APP_TITLE)
        self.geometry("1180x740")
        self.minsize(980, 680)
        self.configure(bg=COLORS["ink"])
        self.option_add("*tearOff", False)

        if sys.platform == "win32":
            try:
                import ctypes

                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
            except Exception:
                pass

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "Gold.Horizontal.TProgressbar",
            background=COLORS["gold"],
            troughcolor=COLORS["paper_soft"],
            bordercolor=COLORS["paper_soft"],
            lightcolor=COLORS["gold"],
            darkcolor=COLORS["gold"],
        )

    def _configure_window_icon(self) -> None:
        icon_png = resource_path("assets", "extrator_precatorios.png")
        icon_ico = resource_path("assets", "extrator_precatorios.ico")

        if icon_png.exists():
            self._window_icon_photo = tk.PhotoImage(file=str(icon_png))
            self.iconphoto(True, self._window_icon_photo)
            self._brand_photo = self._window_icon_photo.subsample(8, 8)

        if icon_ico.exists():
            try:
                self.iconbitmap(default=str(icon_ico))
            except tk.TclError:
                pass

    def _build_layout(self) -> None:
        shell = tk.Frame(self, bg=COLORS["ink"])
        shell.pack(fill="both", expand=True, padx=18, pady=18)

        rail = tk.Frame(shell, bg=COLORS["ink_soft"], width=324, padx=28, pady=26)
        rail.pack(side="left", fill="y")
        rail.pack_propagate(False)

        stage = tk.Frame(shell, bg=COLORS["paper"], padx=28, pady=26, highlightthickness=1, highlightbackground=COLORS["paper_line"])
        stage.pack(side="left", fill="both", expand=True)

        self._build_brand_rail(rail)
        self._build_workspace(stage)

    def _build_brand_rail(self, parent: tk.Frame) -> None:
        brand_top = tk.Frame(parent, bg=COLORS["ink_soft"])
        brand_top.pack(fill="x")

        if self._brand_photo is not None:
            tk.Label(brand_top, image=self._brand_photo, bg=COLORS["ink_soft"]).pack(anchor="w", pady=(0, 14))

        tk.Label(
            brand_top,
            text="Extrator de\nPrecatórios",
            font=("Georgia", 27, "bold"),
            fg=COLORS["paper"],
            bg=COLORS["ink_soft"],
            justify="left",
        ).pack(anchor="w")

        tk.Label(
            brand_top,
            text="Leitura institucional de PDF com consolidação auditável, regras de cessão e validação cruzada em Excel.",
            font=FONTS["body_small"],
            fg=COLORS["gold_soft"],
            bg=COLORS["ink_soft"],
            justify="left",
            wraplength=250,
        ).pack(anchor="w", pady=(10, 18))

        tk.Frame(parent, bg=COLORS["gold"], height=2).pack(fill="x", pady=(0, 18))

        self._build_side_section(
            parent,
            "Fluxo curado",
            [
                ("1", "PDF tratado sem repetir cabeçalhos e sem quebrar o registro."),
                ("2", "Regras de pagamento aplicadas antes de escrever a planilha final."),
                ("3", "Comparativo com base manual e relatório das diferenças confirmadas."),
            ],
        )

        self._build_metric_panel(parent)

        self.open_output_button = tk.Button(
            parent,
            text="Abrir pasta de saída",
            command=self._open_output_dir,
            font=FONTS["button"],
            bg=COLORS["ink_panel"],
            fg=COLORS["paper"],
            activebackground=COLORS["gold"],
            activeforeground=COLORS["ink"],
            bd=0,
            padx=16,
            pady=12,
            cursor="hand2",
        )
        self.open_output_button.pack(fill="x", side="bottom", pady=(18, 0))

    def _build_side_section(self, parent: tk.Frame, title: str, items: list[tuple[str, str]]) -> None:
        section = tk.Frame(parent, bg=COLORS["ink_soft"])
        section.pack(fill="x", pady=(0, 18))

        tk.Label(
            section,
            text=title.upper(),
            font=("Segoe UI Semibold", 9),
            fg=COLORS["gold"],
            bg=COLORS["ink_soft"],
            anchor="w",
        ).pack(anchor="w", pady=(0, 10))

        for number, text in items:
            row = tk.Frame(section, bg=COLORS["ink_panel"], highlightthickness=1, highlightbackground="#30465c", padx=12, pady=11)
            row.pack(fill="x", pady=(0, 8))

            chip = tk.Label(
                row,
                text=number,
                width=3,
                font=("Georgia", 11, "bold"),
                fg=COLORS["ink"],
                bg=COLORS["gold"],
                padx=0,
                pady=6,
            )
            chip.pack(side="left", padx=(0, 10))

            tk.Label(
                row,
                text=text,
                font=FONTS["body_small"],
                fg=COLORS["paper"],
                bg=COLORS["ink_panel"],
                justify="left",
                wraplength=190,
            ).pack(side="left", fill="x", expand=True)

    def _build_metric_panel(self, parent: tk.Frame) -> None:
        wrapper = tk.Frame(parent, bg=COLORS["ink_soft"])
        wrapper.pack(fill="x", pady=(0, 8))

        tk.Label(
            wrapper,
            text="Painel da execução".upper(),
            font=("Segoe UI Semibold", 9),
            fg=COLORS["gold"],
            bg=COLORS["ink_soft"],
        ).pack(anchor="w", pady=(0, 10))

        self._build_metric_card(wrapper, "Arquivo em foco", self.metric_pdf_var).pack(fill="x", pady=(0, 8))
        self._build_metric_card(wrapper, "Registros extraídos", self.metric_records_var).pack(fill="x", pady=(0, 8))
        self._build_metric_card(wrapper, "Validação", self.metric_validation_var).pack(fill="x", pady=(0, 8))
        self._build_metric_card(wrapper, "Qualidade", self.metric_quality_var).pack(fill="x")

    def _build_metric_card(self, parent: tk.Frame, label: str, variable: tk.StringVar) -> tk.Frame:
        card = tk.Frame(parent, bg=COLORS["ink_panel"], highlightthickness=1, highlightbackground="#30465c", padx=12, pady=10)
        tk.Label(
            card,
            text=label.upper(),
            font=("Segoe UI Semibold", 8),
            fg=COLORS["gold_soft"],
            bg=COLORS["ink_panel"],
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            card,
            textvariable=variable,
            font=FONTS["metric"],
            fg=COLORS["paper"],
            bg=COLORS["ink_panel"],
            justify="left",
            wraplength=220,
        ).pack(anchor="w", pady=(6, 0))
        return card

    def _build_workspace(self, parent: tk.Frame) -> None:
        hero = tk.Frame(parent, bg=COLORS["paper"])
        hero.pack(fill="x")

        tk.Label(
            hero,
            text="Consolide o relatório em uma planilha limpa, com regra de pagamento aplicada e trilha de validação pronta.",
            font=FONTS["display"],
            fg=COLORS["text_dark"],
            bg=COLORS["paper"],
            justify="left",
            wraplength=760,
        ).pack(anchor="w")

        tk.Label(
            hero,
            text="Esta interface usa o mesmo núcleo que já foi validado no PDF real: remove cabeçalhos repetidos, preserva o vínculo correto por processo e devolve planilha, relatório técnico e comparativo manual quando solicitado.",
            font=FONTS["body"],
            fg=COLORS["text_soft"],
            bg=COLORS["paper"],
            justify="left",
            wraplength=760,
        ).pack(anchor="w", pady=(10, 16))

        tag_row = tk.Frame(hero, bg=COLORS["paper"])
        tag_row.pack(anchor="w", pady=(0, 18))
        for label in ("PDF", "Regras de cessão", "Excel final", "Validação manual"):
            tk.Label(
                tag_row,
                text=label,
                font=("Segoe UI Semibold", 9),
                fg=COLORS["ink"],
                bg=COLORS["paper_soft"],
                padx=12,
                pady=6,
            ).pack(side="left", padx=(0, 8))

        form_card = tk.Frame(
            parent,
            bg="#fbf8f1",
            padx=22,
            pady=22,
            highlightthickness=1,
            highlightbackground=COLORS["paper_line"],
        )
        form_card.pack(fill="x", pady=(0, 18))

        tk.Label(
            form_card,
            text="Entradas do processamento",
            font=FONTS["display_small"],
            fg=COLORS["text_dark"],
            bg="#fbf8f1",
        ).grid(row=0, column=0, columnspan=3, sticky="w")

        self._build_path_row(
            form_card,
            row=1,
            title="PDF do relatório",
            subtitle="Arquivo base usado para extrair os registros de pagamento.",
            variable=self.pdf_var,
            command=self._choose_pdf,
        )
        self._build_path_row(
            form_card,
            row=2,
            title="Planilha manual",
            subtitle="Opcional. Use quando quiser comparar a saída automática com a base antiga.",
            variable=self.manual_var,
            command=self._choose_manual,
            is_manual=True,
        )
        self._build_path_row(
            form_card,
            row=3,
            title="Pasta de saída",
            subtitle="O app cria os arquivos com timestamp para evitar conflito com planilhas abertas.",
            variable=self.output_var,
            command=self._choose_output_dir,
            folder_mode=True,
        )

        self.validate_check = tk.Checkbutton(
            form_card,
            text="Executar comparação com a planilha manual e gerar relatório comparativo",
            variable=self.validate_var,
            command=self._toggle_manual_state,
            bg="#fbf8f1",
            fg=COLORS["text_dark"],
            activebackground="#fbf8f1",
            activeforeground=COLORS["text_dark"],
            selectcolor=COLORS["paper_soft"],
            font=FONTS["body"],
            anchor="w",
        )
        self.validate_check.grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))

        action_row = tk.Frame(parent, bg=COLORS["paper"])
        action_row.pack(fill="x", pady=(0, 16))

        self.run_button = tk.Button(
            action_row,
            text="Extrair planilha e relatório",
            command=self._start_run,
            font=FONTS["button"],
            bg=COLORS["gold"],
            fg=COLORS["ink"],
            activebackground="#ddb55f",
            activeforeground=COLORS["ink"],
            bd=0,
            padx=18,
            pady=12,
            cursor="hand2",
        )
        self.run_button.pack(side="left")

        self.open_output_stage_button = tk.Button(
            action_row,
            text="Abrir saída",
            command=self._open_output_dir,
            font=FONTS["button"],
            bg=COLORS["ink_soft"],
            fg=COLORS["paper"],
            activebackground=COLORS["ink_panel"],
            activeforeground=COLORS["paper"],
            bd=0,
            padx=16,
            pady=12,
            cursor="hand2",
        )
        self.open_output_stage_button.pack(side="left", padx=(10, 0))

        tk.Label(
            action_row,
            textvariable=self.status_var,
            font=FONTS["body"],
            fg=COLORS["text_soft"],
            bg=COLORS["paper"],
        ).pack(side="left", padx=(14, 0))

        progress_shell = tk.Frame(parent, bg=COLORS["paper"])
        progress_shell.pack(fill="x", pady=(0, 18))

        self.progress = ttk.Progressbar(progress_shell, mode="indeterminate", style="Gold.Horizontal.TProgressbar")
        self.progress.pack(fill="x")

        log_card = tk.Frame(
            parent,
            bg=COLORS["log_bg"],
            padx=18,
            pady=16,
            highlightthickness=1,
            highlightbackground=COLORS["log_line"],
        )
        log_card.pack(fill="both", expand=True)

        tk.Label(
            log_card,
            text="Caderno de execução",
            font=FONTS["display_small"],
            fg=COLORS["paper"],
            bg=COLORS["log_bg"],
        ).pack(anchor="w")
        tk.Label(
            log_card,
            text="Acompanhe o que foi lido, o que foi gerado e qualquer pendência tratada pelo fluxo.",
            font=FONTS["body_small"],
            fg="#9db0c2",
            bg=COLORS["log_bg"],
        ).pack(anchor="w", pady=(4, 12))

        self.log_text = ScrolledText(
            log_card,
            wrap="word",
            height=16,
            font=FONTS["log"],
            bg=COLORS["log_bg"],
            fg="#d9e3ec",
            relief="flat",
            borderwidth=0,
            insertbackground=COLORS["gold"],
            selectbackground=COLORS["ink_panel"],
        )
        self.log_text.pack(fill="both", expand=True)
        self.log_text.configure(state="disabled")

        self._toggle_manual_state()

    def _build_path_row(
        self,
        parent: tk.Frame,
        row: int,
        title: str,
        subtitle: str,
        variable: tk.StringVar,
        command: object,
        is_manual: bool = False,
        folder_mode: bool = False,
    ) -> None:
        base_row = (row - 1) * 2 + 1

        tk.Label(
            parent,
            text=title,
            font=FONTS["title"],
            fg=COLORS["text_dark"],
            bg="#fbf8f1",
        ).grid(row=base_row, column=0, sticky="w", pady=(16, 2))

        tk.Label(
            parent,
            text=subtitle,
            font=FONTS["body_small"],
            fg=COLORS["text_soft"],
            bg="#fbf8f1",
        ).grid(row=base_row + 1, column=0, sticky="w", pady=(0, 8))

        entry = tk.Entry(
            parent,
            textvariable=variable,
            font=FONTS["body"],
            relief="flat",
            bd=0,
            bg="#ffffff",
            fg=COLORS["text_dark"],
            insertbackground=COLORS["text_dark"],
            highlightthickness=1,
            highlightbackground=COLORS["paper_line"],
            highlightcolor=COLORS["gold"],
        )
        entry.grid(row=base_row, column=1, rowspan=2, sticky="ew", padx=(16, 12), pady=(12, 8), ipady=10)

        button_label = "Escolher pasta" if folder_mode else "Escolher"
        button = tk.Button(
            parent,
            text=button_label,
            command=command,
            font=FONTS["button"],
            bg=COLORS["ink_soft"],
            fg=COLORS["paper"],
            activebackground=COLORS["ink_panel"],
            activeforeground=COLORS["paper"],
            bd=0,
            padx=14,
            pady=10,
            cursor="hand2",
        )
        button.grid(row=base_row, column=2, rowspan=2, sticky="nsew", pady=(12, 8))

        parent.columnconfigure(1, weight=1)

        if is_manual:
            self.manual_entry = entry
            self.manual_button = button

    def _toggle_manual_state(self) -> None:
        state = "normal" if self.validate_var.get() else "disabled"
        entry_background = "#ffffff" if state == "normal" else "#ece6d9"

        self.manual_entry.configure(state=state, disabledbackground=entry_background, disabledforeground=COLORS["text_soft"])
        self.manual_button.configure(state=state)

        if self.validate_var.get():
            self.run_button.configure(text="Extrair, validar e gerar relatórios")
            if not self.manual_var.get().strip():
                self.metric_validation_var.set("Base manual pendente")
        else:
            self.run_button.configure(text="Extrair planilha e relatório")
            self.metric_validation_var.set("Validação desativada")

    def _choose_pdf(self) -> None:
        selected = filedialog.askopenfilename(
            title="Escolha o PDF do relatório",
            filetypes=[("Arquivos PDF", "*.pdf")],
            initialdir=str((self.root_dir / "entrada").resolve()),
        )
        if selected:
            self.pdf_var.set(selected)
            self.metric_pdf_var.set(Path(selected).name)

    def _choose_manual(self) -> None:
        initial_dir = self.validation_dir if self.validation_dir is not None else self.root_dir
        selected = filedialog.askopenfilename(
            title="Escolha a planilha manual",
            filetypes=[("Planilhas Excel", "*.xlsx")],
            initialdir=str(initial_dir.resolve()),
        )
        if selected:
            self.manual_var.set(selected)
            self.validate_var.set(True)
            self._toggle_manual_state()

    def _choose_output_dir(self) -> None:
        selected = filedialog.askdirectory(
            title="Escolha a pasta de saída",
            initialdir=self.output_var.get() or str(self.root_dir),
        )
        if selected:
            self.output_var.set(selected)
            self.last_output_dir = Path(selected)

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _open_output_dir(self) -> None:
        target = Path(self.output_var.get().strip()) if self.output_var.get().strip() else self.last_output_dir
        target.mkdir(parents=True, exist_ok=True)
        open_path(target)

    def _start_run(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            return

        pdf_path = Path(self.pdf_var.get().strip())
        output_dir = Path(self.output_var.get().strip())
        manual_path = Path(self.manual_var.get().strip()) if self.manual_var.get().strip() else None

        if not pdf_path.exists():
            messagebox.showerror(APP_TITLE, "Escolha um PDF válido antes de continuar.")
            return
        if self.validate_var.get() and (manual_path is None or not manual_path.exists()):
            messagebox.showerror(APP_TITLE, "Escolha uma planilha manual válida ou desative a comparação.")
            return

        output_dir.mkdir(parents=True, exist_ok=True)
        self.last_output_dir = output_dir

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        self.metric_pdf_var.set(pdf_path.name)
        self.metric_records_var.set("Lendo o PDF...")
        self.metric_validation_var.set("Aguardando comparação" if self.validate_var.get() else "Validação desativada")
        self.metric_quality_var.set("Em processamento")

        self.run_button.configure(state="disabled")
        self.status_var.set("Processando o relatório e montando a saída...")
        self.progress.start(10)

        self.worker_thread = threading.Thread(
            target=self._run_pipeline,
            args=(pdf_path, output_dir, manual_path, self.validate_var.get()),
            daemon=True,
        )
        self.worker_thread.start()

    def _run_pipeline(
        self,
        pdf_path: Path,
        output_dir: Path,
        manual_path: Path | None,
        should_validate: bool,
    ) -> None:
        stamp = timestamp_suffix()
        excel_path = output_dir / f"precatorios_extraidos_{stamp}.xlsx"
        extract_report_path = output_dir / f"precatorios_validacao_{stamp}.txt"
        compare_excel_path = output_dir / f"validacao_comparativa_{stamp}.xlsx"
        compare_report_path = output_dir / f"validacao_comparativa_{stamp}.txt"

        try:
            self.log_queue.put(("log", f"PDF selecionado: {pdf_path}"))
            self.log_queue.put(("log", "Executando extração do PDF..."))
            extraction_result = run_extraction(pdf_path, excel_path, extract_report_path)

            unresolved = sum(1 for item in extraction_result.built_rows if item.warnings)
            self.log_queue.put(("metrics", {"records": str(len(extraction_result.built_rows)), "quality": f"{unresolved} pendências" if unresolved else "Leitura limpa"}))
            self.log_queue.put(("log", f"Registros extraídos: {len(extraction_result.built_rows)}"))
            self.log_queue.put(("log", f"Pendências internas: {unresolved}"))
            self.log_queue.put(("log", f"Planilha criada: {excel_path.name}"))
            self.log_queue.put(("log", f"Relatório técnico: {extract_report_path.name}"))

            validation_result = None
            if should_validate and manual_path is not None:
                self.log_queue.put(("log", "Executando comparação com a planilha manual..."))
                validation_result = run_validation(
                    manual_source=manual_path,
                    auto_path=excel_path,
                    pdf_path=pdf_path,
                    output_path=compare_excel_path,
                    report_path=compare_report_path,
                )
                summary = validation_result.summary
                self.log_queue.put(
                    (
                        "metrics",
                        {
                            "validation": f"{summary['matches_exatos']} matches exatos",
                            "quality": f"{summary['matches_financeiros_nome_diferente']} nome | {summary['matches_processo_nome_valor_diferente']} valor",
                        },
                    )
                )
                self.log_queue.put(("log", f"Aba manual usada: {validation_result.sheet_name}"))
                self.log_queue.put(("log", f"Matches exatos: {summary['matches_exatos']}"))
                self.log_queue.put(("log", f"Nome divergente: {summary['matches_financeiros_nome_diferente']}"))
                self.log_queue.put(("log", f"Valor/metadado divergente: {summary['matches_processo_nome_valor_diferente']}"))
                self.log_queue.put(("log", f"Só na automação: {summary['somente_automatica']}"))
                self.log_queue.put(("log", f"Comparativo Excel: {compare_excel_path.name}"))
                self.log_queue.put(("log", f"Comparativo texto: {compare_report_path.name}"))

            self.log_queue.put(
                (
                    "done",
                    {
                        "excel_path": excel_path,
                        "extract_report_path": extract_report_path,
                        "validation_result": validation_result,
                    },
                )
            )
        except Exception:
            self.log_queue.put(("error", traceback.format_exc()))

    def _apply_metrics(self, payload: dict[str, str]) -> None:
        if "pdf" in payload:
            self.metric_pdf_var.set(payload["pdf"])
        if "records" in payload:
            self.metric_records_var.set(payload["records"])
        if "validation" in payload:
            self.metric_validation_var.set(payload["validation"])
        if "quality" in payload:
            self.metric_quality_var.set(payload["quality"])

    def _drain_queue(self) -> None:
        try:
            while True:
                event, payload = self.log_queue.get_nowait()
                if event == "log":
                    self._append_log(str(payload))
                elif event == "metrics":
                    self._apply_metrics(payload)
                elif event == "done":
                    self._handle_success(payload)
                elif event == "error":
                    self._handle_error(str(payload))
        except queue.Empty:
            pass
        finally:
            self.after(150, self._drain_queue)

    def _handle_success(self, payload: dict[str, object]) -> None:
        self.progress.stop()
        self.run_button.configure(state="normal")
        self.status_var.set("Processamento concluído com sucesso.")

        lines = [
            f"Planilha extraída: {payload['excel_path']}",
            f"Relatório técnico: {payload['extract_report_path']}",
        ]

        validation_result = payload.get("validation_result")
        if validation_result is not None:
            summary = validation_result.summary
            lines.extend(
                [
                    f"Comparativo Excel: {validation_result.output_path}",
                    f"Comparativo texto: {validation_result.report_path}",
                    f"Matches exatos: {summary['matches_exatos']}",
                    f"Nome divergente: {summary['matches_financeiros_nome_diferente']}",
                    f"Valor/metadado divergente: {summary['matches_processo_nome_valor_diferente']}",
                    f"Só na automação: {summary['somente_automatica']}",
                ]
            )
        else:
            self.metric_validation_var.set("Validação não executada")

        messagebox.showinfo(APP_TITLE, "\n".join(lines))

    def _handle_error(self, stacktrace: str) -> None:
        self.progress.stop()
        self.run_button.configure(state="normal")
        self.status_var.set("Falha na execução.")
        self.metric_quality_var.set("Falha")
        self._append_log(stacktrace)
        messagebox.showerror(APP_TITLE, "O processamento falhou. Veja o log para os detalhes.")


if __name__ == "__main__":
    app = ExtractorApp()
    app.mainloop()
