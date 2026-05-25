from __future__ import annotations

import platform
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .cli import _effective_backend
from .excel_com import create_workbook as create_workbook_com
from .excel_com import install_client as install_client_com
from .openpyxl_backend import create_workbook_ui, install_client_ui_file

EXCEL_GREEN = "#217346"
EXCEL_DARK = "#185C37"
EXCEL_LIGHT = "#E7F3EC"
EXCEL_BORDER = "#C8DCCF"
TEXT_DARK = "#1F2933"
TEXT_MUTED = "#5F6B7A"
APP_ID = "DAChernikov.TrinoExcelClient.0.1.0"


def _default_output_dir() -> Path:
    desktop = Path.home() / "Desktop"
    return desktop if desktop.exists() else Path.home()


def _desktop_path(filename: str) -> Path:
    return _default_output_dir() / filename


def _resource_path(relative_path: str) -> Path:
    bundle_root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[2]))
    return bundle_root / relative_path


def _set_windows_app_id() -> None:
    if platform.system() != "Windows":
        return

    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception:
        pass


class TrinoExcelClientApp(tk.Tk):
    def __init__(self, dry_run: bool = False) -> None:
        _set_windows_app_id()
        super().__init__()
        self.title("Trino Excel Client")
        self.geometry("980x680")
        self.minsize(900, 620)
        self.configure(bg="#F7FAF8")
        self.events: queue.Queue[tuple[str, str]] = queue.Queue()

        self.create_backend = tk.StringVar(value="auto")
        self.install_backend = tk.StringVar(value="auto")
        self.sheet_prefix = tk.StringVar(value="Trino")
        self.create_output = tk.StringVar(value=str(_desktop_path("Trino_REST_Client.xlsx")))
        self.install_source = tk.StringVar(value="")
        self.install_output = tk.StringVar(value=str(_desktop_path("Workbook_With_Trino.xlsx")))
        self.overwrite = tk.BooleanVar(value=True)
        self.visible_excel = tk.BooleanVar(value=False)
        self.dry_run = tk.BooleanVar(value=dry_run)

        self._set_app_icon()
        self._configure_style()
        self._build_layout()
        self.after(100, self._drain_events)

    def _set_app_icon(self) -> None:
        icon_path = _resource_path("assets/app_icon.ico")
        if not icon_path.exists():
            return

        try:
            self.iconbitmap(default=str(icon_path))
        except tk.TclError:
            pass

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TFrame", background="#F7FAF8")
        style.configure("Surface.TFrame", background="white", relief="flat")
        style.configure("TLabel", background="#F7FAF8", foreground=TEXT_DARK, font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background="#F7FAF8", foreground=TEXT_MUTED, font=("Segoe UI", 9))
        style.configure("Title.TLabel", background="#F7FAF8", foreground=EXCEL_DARK, font=("Segoe UI", 22, "bold"))
        style.configure("Subtitle.TLabel", background="#F7FAF8", foreground=TEXT_MUTED, font=("Segoe UI", 10))
        style.configure("CardTitle.TLabel", background="white", foreground=EXCEL_DARK, font=("Segoe UI", 13, "bold"))
        style.configure("Card.TLabel", background="white", foreground=TEXT_DARK, font=("Segoe UI", 10))
        style.configure("CardMuted.TLabel", background="white", foreground=TEXT_MUTED, font=("Segoe UI", 9))
        style.configure("TButton", font=("Segoe UI", 10), padding=(12, 8))
        style.configure("Accent.TButton", background=EXCEL_GREEN, foreground="white", font=("Segoe UI", 10, "bold"))
        style.map("Accent.TButton", background=[("active", EXCEL_DARK)], foreground=[("active", "white")])
        style.configure("TCheckbutton", background="#F7FAF8", foreground=TEXT_DARK, font=("Segoe UI", 10))
        style.configure("TNotebook", background="#F7FAF8", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(16, 10), font=("Segoe UI", 10))

    def _build_layout(self) -> None:
        outer = ttk.Frame(self, padding=24)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 18))
        ttk.Label(header, text="Trino Excel Client", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Создание Excel-клиента для Trino REST API и встраивание клиента в существующие книги.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        controls = ttk.Frame(outer)
        controls.pack(fill="x", pady=(0, 12))
        ttk.Checkbutton(controls, text="Dry-run режим", variable=self.dry_run).pack(side="left")
        ttk.Checkbutton(controls, text="Показывать Excel при COM-операциях", variable=self.visible_excel).pack(side="left", padx=(18, 0))
        ttk.Checkbutton(controls, text="Перезаписывать файлы", variable=self.overwrite).pack(side="left", padx=(18, 0))

        prefix_frame = ttk.Frame(controls)
        prefix_frame.pack(side="right")
        ttk.Label(prefix_frame, text="Префикс листов").pack(side="left", padx=(0, 8))
        ttk.Entry(prefix_frame, textvariable=self.sheet_prefix, width=18).pack(side="left")

        notebook = ttk.Notebook(outer)
        notebook.pack(fill="both", expand=True)
        notebook.add(self._build_create_tab(notebook), text="Создать шаблон")
        notebook.add(self._build_install_tab(notebook), text="Встроить в книгу")
        notebook.add(self._build_notes_tab(notebook), text="Подсказки")

        log_frame = ttk.Frame(outer, padding=(0, 14, 0, 0))
        log_frame.pack(fill="both")
        ttk.Label(log_frame, text="Журнал операций", style="Muted.TLabel").pack(anchor="w")
        self.log = tk.Text(log_frame, height=8, bg="#102018", fg="#DDF7E7", insertbackground="white", relief="flat", padx=12, pady=10)
        self.log.pack(fill="both", expand=False, pady=(6, 0))
        self._log("Готово. На macOS используйте backend openpyxl или dry-run для COM-сценариев.")

    def _card(self, parent) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Surface.TFrame", padding=20)
        frame.pack(fill="x", padx=2, pady=2)
        return frame

    def _build_create_tab(self, parent) -> ttk.Frame:
        tab = ttk.Frame(parent, padding=18)
        card = self._card(tab)
        ttk.Label(card, text="Новая Excel-книга", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            card,
            text="На Windows backend auto создаст UI и встроит Power Query. На macOS auto создаст UI-preview без Power Query.",
            style="CardMuted.TLabel",
        ).pack(anchor="w", pady=(4, 18))

        self._path_row(card, "Файл результата", self.create_output, save=True)
        self._backend_row(card, self.create_backend)

        ttk.Button(card, text="Создать книгу", style="Accent.TButton", command=self._create_workbook).pack(anchor="e", pady=(18, 0))
        return tab

    def _build_install_tab(self, parent) -> ttk.Frame:
        tab = ttk.Frame(parent, padding=18)
        card = self._card(tab)
        ttk.Label(card, text="Встраивание клиента", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            card,
            text="Существующие листы сохраняются. Клиент добавляет Trino Config, Trino Query, Trino Result и Trino Help.",
            style="CardMuted.TLabel",
        ).pack(anchor="w", pady=(4, 18))

        self._path_row(card, "Исходная книга", self.install_source, save=False)
        self._path_row(card, "Файл результата", self.install_output, save=True)
        self._backend_row(card, self.install_backend)

        ttk.Button(card, text="Встроить клиент", style="Accent.TButton", command=self._install_client).pack(anchor="e", pady=(18, 0))
        return tab

    def _build_notes_tab(self, parent) -> ttk.Frame:
        tab = ttk.Frame(parent, padding=18)
        card = self._card(tab)
        ttk.Label(card, text="Рабочие ограничения", style="CardTitle.TLabel").pack(anchor="w")
        notes = [
            "Power Query встраивается только на Windows через desktop Excel COM.",
            "На macOS можно проверять UI, dry-run сценарии и создавать preview-книги через openpyxl.",
            "По умолчанию SQL оборачивается внешним LIMIT 100000.",
            "Результат всегда загружается на отдельный лист Trino Result, чтобы refresh не менял Trino Query.",
            "Для больших выгрузок используйте фильтры, агрегаты и внешние export pipeline, а не Excel sheet.",
        ]
        for note in notes:
            ttk.Label(card, text=f"- {note}", style="Card.TLabel").pack(anchor="w", pady=4)
        return tab

    def _path_row(self, parent, label: str, variable: tk.StringVar, save: bool) -> None:
        row = ttk.Frame(parent, style="Surface.TFrame")
        row.pack(fill="x", pady=7)
        ttk.Label(row, text=label, style="Card.TLabel", width=18).pack(side="left")
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True, padx=(8, 8))
        command = lambda: self._choose_path(variable, save=save)
        ttk.Button(row, text="Выбрать", command=command).pack(side="left")

    def _backend_row(self, parent, variable: tk.StringVar) -> None:
        row = ttk.Frame(parent, style="Surface.TFrame")
        row.pack(fill="x", pady=7)
        ttk.Label(row, text="Backend", style="Card.TLabel", width=18).pack(side="left")
        combo = ttk.Combobox(row, values=("auto", "com", "openpyxl"), textvariable=variable, width=16, state="readonly")
        combo.pack(side="left", padx=(8, 0))
        ttk.Label(row, text=f"Текущая платформа: {platform.system()}", style="CardMuted.TLabel").pack(side="left", padx=(18, 0))

    def _choose_path(self, variable: tk.StringVar, save: bool) -> None:
        current = Path(variable.get()).expanduser()
        initial_dir = current.parent if current.parent.exists() else _default_output_dir()
        initial_file = current.name

        if save:
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialdir=str(initial_dir),
                initialfile=initial_file,
                filetypes=(("Excel workbook", "*.xlsx"), ("Excel macro workbook", "*.xlsm"), ("All files", "*.*")),
            )
        else:
            path = filedialog.askopenfilename(
                initialdir=str(initial_dir),
                filetypes=(("Excel workbook", "*.xlsx *.xlsm"), ("All files", "*.*")),
            )
        if path:
            variable.set(path)

    def _create_workbook(self) -> None:
        output = Path(self.create_output.get()).expanduser()
        backend = _effective_backend(self.create_backend.get())
        self._run_async("create", self._create_workbook_impl, output, backend)

    def _create_workbook_impl(self, output: Path, backend: str) -> str:
        if self.dry_run.get():
            return (
                f"DRY-RUN create: backend={backend}, output={output}, "
                f"overwrite={self.overwrite.get()}, sheet_prefix={self.sheet_prefix.get()}"
            )
        if backend == "com":
            result = create_workbook_com(
                output_path=output,
                overwrite=self.overwrite.get(),
                visible=self.visible_excel.get(),
                sheet_prefix=self.sheet_prefix.get(),
            )
            return f"Создана рабочая COM-книга: {result}"
        result = create_workbook_ui(
            output_path=output,
            overwrite=self.overwrite.get(),
            sheet_prefix=self.sheet_prefix.get(),
        )
        return f"Создана UI-preview книга: {result}"

    def _install_client(self) -> None:
        source = Path(self.install_source.get()).expanduser()
        output_text = self.install_output.get().strip()
        output = Path(output_text).expanduser() if output_text else None
        backend = _effective_backend(self.install_backend.get())
        self._run_async("install", self._install_client_impl, source, output, backend)

    def _install_client_impl(self, source: Path, output: Path | None, backend: str) -> str:
        if self.dry_run.get():
            return (
                f"DRY-RUN install: backend={backend}, source={source}, output={output}, "
                f"overwrite={self.overwrite.get()}, sheet_prefix={self.sheet_prefix.get()}"
            )
        if not source.exists():
            raise FileNotFoundError(f"Исходная книга не найдена: {source}")
        if backend == "com":
            result = install_client_com(
                workbook_path=source,
                output_path=output,
                overwrite=self.overwrite.get(),
                visible=self.visible_excel.get(),
                sheet_prefix=self.sheet_prefix.get(),
            )
            return f"Клиент встроен через COM: {result}"
        result = install_client_ui_file(
            workbook_path=source,
            output_path=output,
            overwrite=self.overwrite.get(),
            sheet_prefix=self.sheet_prefix.get(),
        )
        return f"UI-клиент встроен: {result}"

    def _run_async(self, action: str, func, *args) -> None:
        self._log(f"Запуск: {action}")

        def worker() -> None:
            try:
                result = func(*args)
            except Exception as exc:
                self.events.put(("error", str(exc)))
            else:
                self.events.put(("info", result))

        threading.Thread(target=worker, daemon=True).start()

    def _drain_events(self) -> None:
        while True:
            try:
                kind, text = self.events.get_nowait()
            except queue.Empty:
                break
            self._log(text)
            if kind == "error":
                messagebox.showerror("Ошибка", text)
        self.after(100, self._drain_events)

    def _log(self, message: str) -> None:
        self.log.insert("end", message + "\n")
        self.log.see("end")


def run_app(dry_run: bool = False) -> None:
    app = TrinoExcelClientApp(dry_run=dry_run)
    app.mainloop()


def main() -> None:
    run_app(dry_run=False)


if __name__ == "__main__":
    main()
