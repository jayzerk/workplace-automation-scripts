"""Windows desktop launcher for the project's Excel tools."""

from __future__ import annotations

import os
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from Scripts.excel_tools import consolidate_cbt, split_excel


def application_dir() -> Path:
    """Return the executable folder when frozen, or the project folder in source."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


class ExcelToolsApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Tools")
        self.geometry("780x520")
        self.minsize(700, 470)

        self.events: queue.Queue[tuple] = queue.Queue()
        self.running = False
        self.current_output_dir: Path | None = None

        self._configure_style()
        self._build_ui()
        self.after(100, self._process_events)

    def _configure_style(self):
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"))
        style.configure("Heading.TLabel", font=("Segoe UI", 12, "bold"))
        style.configure("TButton", padding=(10, 7))
        style.configure("Tool.TButton", anchor="w", padding=(14, 12))

    def _build_ui(self):
        root = ttk.Frame(self, padding=20)
        root.pack(fill="both", expand=True)

        ttk.Label(root, text="Excel Tools", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            root,
            text="Choose a tool, select its input, and run it from one window.",
        ).pack(anchor="w", pady=(2, 16))

        body = ttk.Frame(root)
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        sidebar = ttk.Frame(body, padding=(0, 0, 18, 0))
        sidebar.grid(row=0, column=0, sticky="ns")
        ttk.Button(
            sidebar,
            text="Excel Splitter",
            style="Tool.TButton",
            width=20,
            command=lambda: self._show_tool("splitter"),
        ).pack(fill="x", pady=(0, 6))
        ttk.Button(
            sidebar,
            text="CBT Consolidation",
            style="Tool.TButton",
            width=20,
            command=lambda: self._show_tool("cbt"),
        ).pack(fill="x")

        self.content = ttk.Frame(body, padding=18, relief="solid", borderwidth=1)
        self.content.grid(row=0, column=1, sticky="nsew")

        status_frame = ttk.Frame(root)
        status_frame.pack(fill="x", pady=(16, 0))
        self.progress = ttk.Progressbar(status_frame, mode="determinate")
        self.progress.pack(fill="x")
        self.status = tk.StringVar(value="Ready")
        ttk.Label(status_frame, textvariable=self.status).pack(anchor="w", pady=(5, 0))

        self.open_output_button = ttk.Button(
            status_frame,
            text="Open Output Folder",
            command=self._open_output_folder,
            state="disabled",
        )
        self.open_output_button.pack(anchor="e", pady=(6, 0))

        base = application_dir()
        self.split_input = tk.StringVar()
        self.split_output = tk.StringVar(
            value=str(base / "Output" / "ExcelSplitter")
        )
        self.split_rows = tk.StringVar(value="20000")
        self.cbt_input = tk.StringVar()
        self.cbt_output = tk.StringVar(value=str(base / "Output" / "CBT"))
        self.cbt_sheet = tk.StringVar(value="Sheet2")

        self._show_tool("splitter")

    def _show_tool(self, tool: str):
        if self.running:
            return
        for child in self.content.winfo_children():
            child.destroy()
        self.content.columnconfigure(1, weight=1)

        if tool == "splitter":
            self._build_splitter_form()
        else:
            self._build_cbt_form()

    def _build_splitter_form(self):
        ttk.Label(
            self.content, text="Excel Splitter", style="Heading.TLabel"
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 16))
        self._file_row(1, "Input workbook", self.split_input)
        self._folder_row(2, "Output folder", self.split_output)
        ttk.Label(self.content, text="Rows per file").grid(
            row=3, column=0, sticky="w", pady=8
        )
        ttk.Entry(self.content, textvariable=self.split_rows, width=18).grid(
            row=3, column=1, sticky="w", padx=(12, 8), pady=8
        )
        ttk.Button(
            self.content, text="Split Workbook", command=self._start_split
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(18, 0))

    def _build_cbt_form(self):
        ttk.Label(
            self.content, text="CBT Consolidation", style="Heading.TLabel"
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 16))
        self._file_row(1, "Input workbook", self.cbt_input)
        self._folder_row(2, "Output folder", self.cbt_output)
        ttk.Label(self.content, text="Worksheet name").grid(
            row=3, column=0, sticky="w", pady=8
        )
        ttk.Entry(self.content, textvariable=self.cbt_sheet, width=24).grid(
            row=3, column=1, sticky="w", padx=(12, 8), pady=8
        )
        ttk.Button(
            self.content, text="Create Report", command=self._start_cbt
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(18, 0))

    def _file_row(self, row: int, label: str, variable: tk.StringVar):
        ttk.Label(self.content, text=label).grid(row=row, column=0, sticky="w", pady=8)
        ttk.Entry(self.content, textvariable=variable).grid(
            row=row, column=1, sticky="ew", padx=(12, 8), pady=8
        )
        ttk.Button(
            self.content,
            text="Browse...",
            command=lambda: self._choose_file(variable),
        ).grid(row=row, column=2, pady=8)

    def _folder_row(self, row: int, label: str, variable: tk.StringVar):
        ttk.Label(self.content, text=label).grid(row=row, column=0, sticky="w", pady=8)
        ttk.Entry(self.content, textvariable=variable).grid(
            row=row, column=1, sticky="ew", padx=(12, 8), pady=8
        )
        ttk.Button(
            self.content,
            text="Browse...",
            command=lambda: self._choose_folder(variable),
        ).grid(row=row, column=2, pady=8)

    def _choose_file(self, variable: tk.StringVar):
        filename = filedialog.askopenfilename(
            title="Select Excel workbook",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if filename:
            variable.set(filename)

    def _choose_folder(self, variable: tk.StringVar):
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            variable.set(folder)

    def _start_split(self):
        try:
            input_file = self._validate_input_file(self.split_input.get())
            output_dir = self._validate_output_dir(self.split_output.get())
        except (ValueError, FileNotFoundError, NotADirectoryError) as error:
            messagebox.showerror("Invalid selection", str(error))
            return

        try:
            rows_per_file = int(self.split_rows.get().strip())
            if rows_per_file <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Invalid rows", "Rows per file must be a whole number greater than zero."
            )
            return

        self._set_running("Reading workbook...")
        self.progress.configure(mode="determinate", maximum=100, value=0)

        def work():
            try:
                def progress(current, total, output_file):
                    self.events.put(
                        (
                            "progress",
                            current / total * 100,
                            f"Created {current} of {total}: {output_file.name}",
                        )
                    )

                files = split_excel(
                    input_file, output_dir, rows_per_file, progress_callback=progress
                )
                self.events.put(
                    (
                        "success",
                        output_dir,
                        f"Created {len(files)} split workbook(s).",
                    )
                )
            except Exception as error:
                self.events.put(("error", error))

        threading.Thread(target=work, daemon=True).start()

    def _start_cbt(self):
        try:
            input_file = self._validate_input_file(self.cbt_input.get())
            output_dir = self._validate_output_dir(self.cbt_output.get())
            sheet_name = self.cbt_sheet.get().strip()
            if not sheet_name:
                raise ValueError("Worksheet name cannot be blank.")
        except (ValueError, FileNotFoundError, NotADirectoryError) as error:
            messagebox.showerror("Invalid selection", str(error))
            return

        self._set_running("Creating consolidated report...")
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)

        def work():
            try:
                output_file = consolidate_cbt(input_file, output_dir, sheet_name)
                self.events.put(
                    ("success", output_dir, f"Report created: {output_file.name}")
                )
            except Exception as error:
                self.events.put(("error", error))

        threading.Thread(target=work, daemon=True).start()

    @staticmethod
    def _validate_input_file(value: str) -> Path:
        path = Path(value.strip())
        if not path.is_file():
            raise FileNotFoundError("Please select an existing Excel workbook.")
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("Please select an .xlsx or .xlsm workbook.")
        return path

    @staticmethod
    def _validate_output_dir(value: str) -> Path:
        if not value.strip():
            raise NotADirectoryError("Please select an output folder.")
        path = Path(value.strip())
        path.mkdir(parents=True, exist_ok=True)
        if not path.is_dir():
            raise NotADirectoryError("The selected output location is not a folder.")
        return path

    def _set_running(self, message: str):
        self.running = True
        self.status.set(message)
        self.open_output_button.configure(state="disabled")
        self._set_content_state("disabled")

    def _set_content_state(self, state: str):
        for child in self.content.winfo_children():
            if child.winfo_class() in {"TButton", "TEntry"}:
                child.configure(state=state)

    def _process_events(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "progress":
                    _, value, message = event
                    self.progress.configure(value=value)
                    self.status.set(message)
                elif event[0] == "success":
                    _, output_dir, message = event
                    self._finish_run(message, output_dir)
                    messagebox.showinfo("Completed", message)
                elif event[0] == "error":
                    self._finish_run("Operation failed.")
                    messagebox.showerror("Error", str(event[1]))
        except queue.Empty:
            pass
        self.after(100, self._process_events)

    def _finish_run(self, message: str, output_dir: Path | None = None):
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.running = False
        self.status.set(message)
        self._set_content_state("normal")
        self.current_output_dir = output_dir
        if output_dir is not None:
            self.open_output_button.configure(state="normal")

    def _open_output_folder(self):
        if self.current_output_dir and self.current_output_dir.is_dir():
            os.startfile(self.current_output_dir)


if __name__ == "__main__":
    ExcelToolsApp().mainloop()
