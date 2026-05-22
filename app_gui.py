from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from PIL import Image, ImageTk

from leitor_pdf import count_auto_adjustment_rows, count_manual_review_rows, run_extraction


APP_TITLE = "Extrator de Precatórios"
APP_ID = "gabriel.bispo.extrator.precatorios"
BRAND_ICON_SIZE = 118
WINDOW_ICON_SIZES = (256, 128, 64, 48, 32, 16)

COLORS = {
    "ink": "#101b25",
    "ink_soft": "#182635",
    "ink_panel": "#1e3143",
    "paper": "#f5efe4",
    "paper_soft": "#efe3cc",
    "paper_panel": "#fbf7ef",
    "paper_line": "#d8c7aa",
    "text_dark": "#152230",
    "text_soft": "#5b6a78",
    "gold": "#c79734",
    "gold_soft": "#f0d79d",
    "success": "#2c7a5a",
    "danger": "#a84b3f",
    "log_bg": "#12202d",
    "log_line": "#294155",
}

FONTS = {
    "display": ("Georgia", 24, "bold"),
    "display_small": ("Georgia", 18, "bold"),
    "title": ("Segoe UI Semibold", 11),
    "body": ("Segoe UI", 10),
    "body_small": ("Segoe UI", 9),
    "metric": ("Georgia", 15, "bold"),
    "button": ("Segoe UI Semibold", 10),
    "eyebrow": ("Segoe UI Semibold", 9),
    "log": ("Consolas", 10),
}


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


def enable_high_dpi() -> None:
    if sys.platform != "win32":
        return

    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            import ctypes

            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def resampling_lanczos() -> int:
    if hasattr(Image, "Resampling"):
        return Image.Resampling.LANCZOS
    return Image.LANCZOS


class ExtractorApp(tk.Tk):
    def __init__(self) -> None:
        enable_high_dpi()
        super().__init__()

        self.root_dir = project_root()
        self.log_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker_thread: threading.Thread | None = None

        self.pdf_var = tk.StringVar(value=first_file(self.root_dir / "entrada", "*.pdf"))
        self.output_var = tk.StringVar(value=str((self.root_dir / "saida").resolve()))
        self.status_var = tk.StringVar(value="Pronto para gerar a planilha final.")
        self.metric_pdf_var = tk.StringVar(value="Aguardando PDF")
        self.metric_records_var = tk.StringVar(value="Sem leitura")
        self.metric_quality_var = tk.StringVar(value="Sem execução")
        self.last_output_dir = Path(self.output_var.get())

        self._window_icon_photos: list[object] = []
        self._brand_photo: object | None = None

        self._configure_window()
        self._configure_window_icon()
        self._build_layout()
        self.after(150, self._drain_queue)

    def _configure_window(self) -> None:
        self.title(APP_TITLE)
        self.geometry("1160x720")
        self.minsize(960, 660)
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
            try:
                with Image.open(icon_png) as source:
                    base_image = source.convert("RGBA")
                    sampler = resampling_lanczos()

                    brand_image = base_image.resize((BRAND_ICON_SIZE, BRAND_ICON_SIZE), sampler)
                    self._brand_photo = ImageTk.PhotoImage(brand_image)

                    for size in WINDOW_ICON_SIZES:
                        resized = base_image.resize((size, size), sampler)
                        photo = ImageTk.PhotoImage(resized)
                        self._window_icon_photos.append(photo)

                if self._window_icon_photos:
                    self.iconphoto(True, *self._window_icon_photos)
            except Exception:
                fallback = tk.PhotoImage(file=str(icon_png))
                self._window_icon_photos.append(fallback)
                self.iconphoto(True, fallback)
                self._brand_photo = fallback.subsample(8, 8)

        if icon_ico.exists():
            try:
                self.iconbitmap(default=str(icon_ico))
            except tk.TclError:
                pass

    def _build_layout(self) -> None:
        shell = tk.Frame(self, bg=COLORS["ink"])
        shell.pack(fill="both", expand=True, padx=18, pady=18)

        rail = tk.Frame(shell, bg=COLORS["ink_soft"], width=304, padx=28, pady=28)
        rail.pack(side="left", fill="y")
        rail.pack_propagate(False)

        stage = tk.Frame(
            shell,
            bg=COLORS["paper"],
            padx=32,
            pady=30,
            highlightthickness=1,
            highlightbackground=COLORS["paper_line"],
        )
        stage.pack(side="left", fill="both", expand=True)

        self._build_brand_rail(rail)
        self._build_workspace(stage)

    def _build_brand_rail(self, parent: tk.Frame) -> None:
        brand_top = tk.Frame(parent, bg=COLORS["ink_soft"])
        brand_top.pack(fill="x")

        if self._brand_photo is not None:
            tk.Label(brand_top, image=self._brand_photo, bg=COLORS["ink_soft"]).pack(anchor="w", pady=(0, 16))

        tk.Label(
            brand_top,
            text="Extrator de\nPrecatórios",
            font=("Georgia", 26, "bold"),
            fg=COLORS["paper"],
            bg=COLORS["ink_soft"],
            justify="left",
        ).pack(anchor="w")

        tk.Label(
            brand_top,
            text="Leitura institucional de PDF com consolidação estruturada e saída pronta para operação.",
            font=FONTS["body_small"],
            fg=COLORS["gold_soft"],
            bg=COLORS["ink_soft"],
            justify="left",
            wraplength=228,
        ).pack(anchor="w", pady=(10, 18))

        tk.Frame(parent, bg=COLORS["gold"], height=2).pack(fill="x", pady=(0, 18))

        self._build_side_section(
            parent,
            "Fluxo essencial",
            [
                ("1", "Leitura textual do PDF com limpeza estrutural e agrupamento por processo."),
                ("2", "Aplicação das regras de pagamento antes de montar a planilha final."),
                ("3", "Entrega de Excel consolidado e relatório técnico da execução."),
            ],
        )

        self._build_metric_panel(parent)

    def _build_side_section(self, parent: tk.Frame, title: str, items: list[tuple[str, str]]) -> None:
        section = tk.Frame(parent, bg=COLORS["ink_soft"])
        section.pack(fill="x", pady=(0, 18))

        tk.Label(
            section,
            text=title.upper(),
            font=FONTS["eyebrow"],
            fg=COLORS["gold"],
            bg=COLORS["ink_soft"],
            anchor="w",
        ).pack(anchor="w", pady=(0, 10))

        for number, text in items:
            row = tk.Frame(
                section,
                bg=COLORS["ink_panel"],
                highlightthickness=1,
                highlightbackground="#31475c",
                padx=12,
                pady=11,
            )
            row.pack(fill="x", pady=(0, 8))

            chip = tk.Label(
                row,
                text=number,
                width=3,
                font=("Georgia", 11, "bold"),
                fg=COLORS["ink"],
                bg=COLORS["gold"],
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
                wraplength=184,
            ).pack(side="left", fill="x", expand=True)

    def _build_metric_panel(self, parent: tk.Frame) -> None:
        wrapper = tk.Frame(parent, bg=COLORS["ink_soft"])
        wrapper.pack(fill="x")

        tk.Label(
            wrapper,
            text="Pulso da execução".upper(),
            font=FONTS["eyebrow"],
            fg=COLORS["gold"],
            bg=COLORS["ink_soft"],
        ).pack(anchor="w", pady=(0, 10))

        self._build_metric_card(wrapper, "Arquivo em foco", self.metric_pdf_var).pack(fill="x", pady=(0, 8))
        self._build_metric_card(wrapper, "Registros extraídos", self.metric_records_var).pack(fill="x", pady=(0, 8))
        self._build_metric_card(wrapper, "Qualidade", self.metric_quality_var).pack(fill="x")

    def _build_metric_card(self, parent: tk.Frame, label: str, variable: tk.StringVar) -> tk.Frame:
        card = tk.Frame(
            parent,
            bg=COLORS["ink_panel"],
            highlightthickness=1,
            highlightbackground="#31475c",
            padx=12,
            pady=10,
        )
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
            wraplength=214,
        ).pack(anchor="w", pady=(6, 0))
        return card

    def _build_workspace(self, parent: tk.Frame) -> None:
        hero = tk.Frame(parent, bg=COLORS["paper"])
        hero.pack(fill="x")

        tk.Label(
            hero,
            text="FLUXO PRINCIPAL",
            font=FONTS["eyebrow"],
            fg=COLORS["gold"],
            bg=COLORS["paper"],
        ).pack(anchor="w", pady=(0, 8))

        tk.Label(
            hero,
            text="Transforme o relatório em uma planilha final limpa, com regra de pagamento aplicada e rastreabilidade técnica preservada.",
            font=FONTS["display"],
            fg=COLORS["text_dark"],
            bg=COLORS["paper"],
            justify="left",
            wraplength=760,
        ).pack(anchor="w")

        tk.Label(
            hero,
            text="O núcleo trata o PDF real, remove repetições estruturais e mantém o vínculo correto por processo sem depender de fluxos paralelos.",
            font=FONTS["body"],
            fg=COLORS["text_soft"],
            bg=COLORS["paper"],
            justify="left",
            wraplength=760,
        ).pack(anchor="w", pady=(10, 16))

        tag_row = tk.Frame(hero, bg=COLORS["paper"])
        tag_row.pack(anchor="w", pady=(0, 18))
        for label in ("PDF real", "Regras de cessão", "Excel final", "Relatório técnico"):
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
            bg=COLORS["paper_panel"],
            padx=24,
            pady=24,
            highlightthickness=1,
            highlightbackground=COLORS["paper_line"],
        )
        form_card.pack(fill="x", pady=(0, 18))

        tk.Label(
            form_card,
            text="Entradas do processamento",
            font=FONTS["display_small"],
            fg=COLORS["text_dark"],
            bg=COLORS["paper_panel"],
        ).grid(row=0, column=0, columnspan=3, sticky="w")

        tk.Label(
            form_card,
            text="Fluxo enxuto: informe o relatório e a pasta onde a entrega será gerada.",
            font=FONTS["body_small"],
            fg=COLORS["text_soft"],
            bg=COLORS["paper_panel"],
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 2))

        self._build_path_row(
            form_card,
            row=2,
            title="PDF do relatório",
            subtitle="Arquivo base usado para extrair os registros de pagamento.",
            variable=self.pdf_var,
            command=self._choose_pdf,
        )
        self._build_path_row(
            form_card,
            row=4,
            title="Pasta de saída",
            subtitle="Os arquivos são gerados com timestamp para evitar conflito com versões abertas.",
            variable=self.output_var,
            command=self._choose_output_dir,
            folder_mode=True,
        )

        action_row = tk.Frame(parent, bg=COLORS["paper"])
        action_row.pack(fill="x", pady=(0, 16))

        self.run_button = tk.Button(
            action_row,
            text="Gerar planilha final",
            command=self._start_run,
            font=FONTS["button"],
            bg=COLORS["gold"],
            fg=COLORS["ink"],
            activebackground="#dcb25d",
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
            text="Acompanhe a leitura, os arquivos gerados e qualquer revisão interna sinalizada pelo núcleo.",
            font=FONTS["body_small"],
            fg="#9db0c2",
            bg=COLORS["log_bg"],
        ).pack(anchor="w", pady=(4, 12))

        self.log_text = ScrolledText(
            log_card,
            wrap="word",
            height=15,
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

    def _build_path_row(
        self,
        parent: tk.Frame,
        row: int,
        title: str,
        subtitle: str,
        variable: tk.StringVar,
        command: object,
        folder_mode: bool = False,
    ) -> None:
        tk.Label(
            parent,
            text=title,
            font=FONTS["title"],
            fg=COLORS["text_dark"],
            bg=COLORS["paper_panel"],
        ).grid(row=row, column=0, sticky="w", pady=(16, 2))

        tk.Label(
            parent,
            text=subtitle,
            font=FONTS["body_small"],
            fg=COLORS["text_soft"],
            bg=COLORS["paper_panel"],
        ).grid(row=row + 1, column=0, sticky="w", pady=(0, 8))

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
        entry.grid(row=row, column=1, rowspan=2, sticky="ew", padx=(16, 12), pady=(12, 8), ipady=10)

        button = tk.Button(
            parent,
            text="Escolher pasta" if folder_mode else "Escolher",
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
        button.grid(row=row, column=2, rowspan=2, sticky="nsew", pady=(12, 8))

        parent.columnconfigure(1, weight=1)

    def _choose_pdf(self) -> None:
        selected = filedialog.askopenfilename(
            title="Escolha o PDF do relatório",
            filetypes=[("Arquivos PDF", "*.pdf")],
            initialdir=str((self.root_dir / "entrada").resolve()),
        )
        if selected:
            self.pdf_var.set(selected)
            self.metric_pdf_var.set(Path(selected).name)

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

        if not pdf_path.exists():
            messagebox.showerror(APP_TITLE, "Escolha um PDF válido antes de continuar.")
            return

        output_dir.mkdir(parents=True, exist_ok=True)
        self.last_output_dir = output_dir

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        self.metric_pdf_var.set(pdf_path.name)
        self.metric_records_var.set("Lendo o PDF...")
        self.metric_quality_var.set("Em processamento")

        self.run_button.configure(state="disabled")
        self.status_var.set("Processando o relatório e montando a saída final...")
        self.progress.start(10)

        self.worker_thread = threading.Thread(
            target=self._run_pipeline,
            args=(pdf_path, output_dir),
            daemon=True,
        )
        self.worker_thread.start()

    def _run_pipeline(self, pdf_path: Path, output_dir: Path) -> None:
        stamp = timestamp_suffix()
        excel_path = output_dir / f"precatorios_extraidos_{stamp}.xlsx"
        report_path = output_dir / f"precatorios_relatorio_tecnico_{stamp}.txt"

        try:
            self.log_queue.put(("log", f"PDF selecionado: {pdf_path}"))
            self.log_queue.put(("log", "Executando leitura estruturada do PDF..."))
            extraction_result = run_extraction(pdf_path, excel_path, report_path)

            review_cases = count_manual_review_rows(extraction_result.built_rows)
            auto_adjustments = count_auto_adjustment_rows(extraction_result.built_rows)
            if review_cases:
                quality_label = f"{review_cases} revisão" if review_cases == 1 else f"{review_cases} revisões"
            elif auto_adjustments:
                quality_label = (
                    f"{auto_adjustments} ajuste auto"
                    if auto_adjustments == 1
                    else f"{auto_adjustments} ajustes auto"
                )
            else:
                quality_label = "Leitura limpa"

            self.log_queue.put(
                (
                    "metrics",
                    {
                        "records": str(len(extraction_result.built_rows)),
                        "quality": quality_label,
                    },
                )
            )
            self.log_queue.put(("log", f"Registros extraídos: {len(extraction_result.built_rows)}"))
            self.log_queue.put(("log", f"Revisões manuais sinalizadas: {review_cases}"))
            self.log_queue.put(("log", f"Ajustes automáticos monitorados: {auto_adjustments}"))
            self.log_queue.put(("log", f"Planilha criada: {excel_path.name}"))
            self.log_queue.put(("log", f"Relatório técnico: {report_path.name}"))

            self.log_queue.put(
                (
                    "done",
                    {
                        "excel_path": excel_path,
                        "report_path": report_path,
                        "record_count": len(extraction_result.built_rows),
                        "review_cases": review_cases,
                        "auto_adjustments": auto_adjustments,
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
        self.status_var.set("Entrega concluída com sucesso.")

        lines = [
            f"Planilha extraída: {payload['excel_path']}",
            f"Relatório técnico: {payload['report_path']}",
        ]

        review_cases = int(payload["review_cases"])
        auto_adjustments = int(payload["auto_adjustments"])
        if review_cases:
            lines.append(f"Revisões manuais sinalizadas: {review_cases}")
        if auto_adjustments:
            lines.append(f"Ajustes automáticos monitorados: {auto_adjustments}")
        if not review_cases and not auto_adjustments:
            lines.append("Qualidade da leitura: sem revisões pendentes.")

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
