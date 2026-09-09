"""Persistent three-column queue board for the desktop application."""

from __future__ import annotations

import json
import os
import re
import tkinter as tk
import uuid
import webbrowser
from copy import deepcopy
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, ttk

from Scripts.security_settings import (
    AppSettingsStore,
    get_startup_enabled,
    normalize_allowed_domains,
    set_startup_enabled,
    validate_link,
)


QUEUE_COLUMNS = ("general", "weekly", "monthly")
ARCHIVE_COLUMN = "archive"
ALL_COLUMNS = QUEUE_COLUMNS + (ARCHIVE_COLUMN,)
COLUMN_LABELS = {
    "general": "General",
    "weekly": "Monitored",
    "monthly": "Monthly",
}
DEFAULT_COLOR = "#F2C94C"
COLOR_PATTERN = re.compile(r"^#[0-9A-Fa-f]{6}$")


class QueueStore:
    """Load and atomically save queue bricks in a JSON file."""

    def __init__(self, data_file: str | Path, allowed_domains: list[str] | None = None):
        self.data_file = Path(data_file)
        self.allowed_domains = list(
            allowed_domains if allowed_domains is not None else ["docs.google.com"]
        )
        self.columns: dict[str, list[dict[str, str]]] = {
            column: [] for column in ALL_COLUMNS
        }
        self.archived_during_load = 0

    def load(self) -> None:
        if not self.data_file.exists():
            return

        try:
            payload = json.loads(self.data_file.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as error:
            raise ValueError(f"Could not read queue data: {error}") from error

        raw_columns = payload.get("columns") if isinstance(payload, dict) else None
        if not isinstance(raw_columns, dict):
            raise ValueError("Queue data is invalid: missing 'columns' object.")

        loaded: dict[str, list[dict[str, str]]] = {
            column: [] for column in ALL_COLUMNS
        }
        seen_ids: set[str] = set()
        archived_during_load = 0
        for column in ALL_COLUMNS:
            raw_bricks = raw_columns.get(column, [])
            if not isinstance(raw_bricks, list):
                raise ValueError(f"Queue data is invalid: '{column}' must be a list.")
            for raw_brick in raw_bricks:
                brick = self._validated_brick(raw_brick)
                if brick["id"] in seen_ids:
                    raise ValueError(f"Queue data contains duplicate ID {brick['id']}.")
                seen_ids.add(brick["id"])
                target_column = column
                if column == ARCHIVE_COLUMN:
                    source_column = str(raw_brick.get("source_column", "general"))
                    if source_column not in QUEUE_COLUMNS:
                        source_column = "general"
                    brick["source_column"] = source_column
                    brick["link_error"] = str(
                        raw_brick.get("link_error", "Archived link requires review.")
                    ).strip()
                elif brick["link"]:
                    try:
                        validate_link(brick["link"], self.allowed_domains)
                    except ValueError as error:
                        target_column = ARCHIVE_COLUMN
                        brick["source_column"] = column
                        brick["link_error"] = str(error)
                        archived_during_load += 1
                loaded[target_column].append(brick)
        self.columns = loaded
        self.archived_during_load = archived_during_load

    def _validated_brick(self, raw_brick: object) -> dict[str, str]:
        if not isinstance(raw_brick, dict):
            raise ValueError("Queue data contains a brick that is not an object.")

        identifier = str(raw_brick.get("id", "")).strip()
        name = str(raw_brick.get("name", "")).strip()
        color = str(raw_brick.get("color", "")).strip()
        link = str(raw_brick.get("link", "")).strip()
        if not identifier or not name:
            raise ValueError("Queue data contains a brick without an ID or name.")
        if not COLOR_PATTERN.fullmatch(color):
            raise ValueError(f"Queue brick '{name}' has an invalid color.")
        return {"id": identifier, "name": name, "color": color.upper(), "link": link}

    def save(self) -> None:
        self.data_file.parent.mkdir(parents=True, exist_ok=True)
        temporary_file = self.data_file.with_suffix(self.data_file.suffix + ".tmp")
        payload = {"version": 1, "columns": self.columns}
        try:
            temporary_file.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
            os.replace(temporary_file, self.data_file)
        except OSError:
            try:
                temporary_file.unlink(missing_ok=True)
            except OSError:
                pass
            raise

    def snapshot(self) -> dict[str, list[dict[str, str]]]:
        return deepcopy(self.columns)

    def restore(self, snapshot: dict[str, list[dict[str, str]]]) -> None:
        self.columns = deepcopy(snapshot)

    def add(self, column: str, name: str, color: str, link: str = "") -> str:
        self._validate_column(column)
        brick = {
            "id": uuid.uuid4().hex,
            "name": name.strip(),
            "color": color.upper(),
            "link": validate_link(link, self.allowed_domains),
        }
        if not brick["name"]:
            raise ValueError("Name cannot be blank.")
        if not COLOR_PATTERN.fullmatch(brick["color"]):
            raise ValueError("Color must be a six-digit hexadecimal color.")
        self.columns[column].append(brick)
        return brick["id"]

    def find(self, brick_id: str) -> tuple[str, int, dict[str, str]] | None:
        for column in ALL_COLUMNS:
            for index, brick in enumerate(self.columns[column]):
                if brick["id"] == brick_id:
                    return column, index, brick
        return None

    def update(
        self,
        brick_id: str,
        column: str,
        name: str,
        color: str,
        link: str = "",
    ) -> None:
        self._validate_column(column)
        found = self.find(brick_id)
        if found is None:
            raise KeyError(f"Queue brick not found: {brick_id}")
        old_column, old_index, brick = found
        cleaned_name = name.strip()
        cleaned_color = color.upper()
        if not cleaned_name:
            raise ValueError("Name cannot be blank.")
        if not COLOR_PATTERN.fullmatch(cleaned_color):
            raise ValueError("Color must be a six-digit hexadecimal color.")

        brick.update(
            name=cleaned_name,
            color=cleaned_color,
            link=validate_link(link, self.allowed_domains),
        )
        brick.pop("source_column", None)
        brick.pop("link_error", None)
        if column != old_column:
            self.columns[old_column].pop(old_index)
            self.columns[column].append(brick)

    def delete(self, brick_id: str) -> None:
        found = self.find(brick_id)
        if found is None:
            raise KeyError(f"Queue brick not found: {brick_id}")
        column, index, _ = found
        self.columns[column].pop(index)

    def archive(self, brick_id: str) -> None:
        found = self.find(brick_id)
        if found is None:
            raise KeyError(f"Queue brick not found: {brick_id}")
        column, index, brick = found
        if column == ARCHIVE_COLUMN:
            return
        self.columns[column].pop(index)
        brick["source_column"] = column
        brick["link_error"] = "Archived manually."
        self.columns[ARCHIVE_COLUMN].append(brick)

    def unarchive(self, brick_id: str) -> str:
        found = self.find(brick_id)
        if found is None or found[0] != ARCHIVE_COLUMN:
            raise KeyError(f"Archived queue brick not found: {brick_id}")
        _, index, brick = found
        validate_link(brick["link"], self.allowed_domains)
        target_column = brick.get("source_column", "general")
        if target_column not in QUEUE_COLUMNS:
            target_column = "general"
        self.columns[ARCHIVE_COLUMN].pop(index)
        brick.pop("source_column", None)
        brick.pop("link_error", None)
        self.columns[target_column].append(brick)
        return target_column

    def reorder(self, column: str, brick_id: str, target_index: int) -> bool:
        self._validate_column(column)
        found = self.find(brick_id)
        if found is None or found[0] != column:
            raise KeyError(f"Queue brick not found in {column}: {brick_id}")
        return self.move(brick_id, column, target_index)

    def move(self, brick_id: str, target_column: str, target_index: int) -> bool:
        """Move an active brick within or between active queue columns."""
        self._validate_column(target_column)
        found = self.find(brick_id)
        if found is None or found[0] == ARCHIVE_COLUMN:
            raise KeyError(f"Active queue brick not found: {brick_id}")
        source_column, source_index, brick = found
        self.columns[source_column].pop(source_index)
        target_index = max(0, min(target_index, len(self.columns[target_column])))
        if source_column == target_column and target_index == source_index:
            self.columns[source_column].insert(source_index, brick)
            return False
        self.columns[target_column].insert(target_index, brick)
        return True

    @staticmethod
    def _validate_column(column: str) -> None:
        if column not in QUEUE_COLUMNS:
            raise ValueError(f"Unknown queue column: {column}")


class BrickDialog(tk.Toplevel):
    """Modal dialog used to add or edit a queue brick."""

    def __init__(
        self,
        parent,
        title: str,
        allowed_domains: list[str],
        initial_column: str = "general",
        initial_brick: dict[str, str] | None = None,
    ):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.transient(parent.winfo_toplevel())
        self.result: dict[str, str] | None = None
        self.allowed_domains = allowed_domains

        brick = initial_brick or {}
        self.name_value = tk.StringVar(value=brick.get("name", ""))
        self.column_value = tk.StringVar(value=COLUMN_LABELS[initial_column])
        self.color_value = tk.StringVar(value=brick.get("color", DEFAULT_COLOR))
        self.link_value = tk.StringVar(value=brick.get("link", ""))

        frame = ttk.Frame(self, padding=18)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Name").grid(row=0, column=0, sticky="w", pady=7)
        name_entry = ttk.Entry(frame, textvariable=self.name_value, width=44)
        name_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=(12, 0), pady=7)

        ttk.Label(frame, text="Column").grid(row=1, column=0, sticky="w", pady=7)
        column_picker = ttk.Combobox(
            frame,
            textvariable=self.column_value,
            values=[COLUMN_LABELS[column] for column in QUEUE_COLUMNS],
            state="readonly",
            width=20,
        )
        column_picker.grid(row=1, column=1, columnspan=2, sticky="w", padx=(12, 0), pady=7)

        ttk.Label(frame, text="Color").grid(row=2, column=0, sticky="w", pady=7)
        self.color_preview = tk.Label(
            frame,
            width=8,
            relief="sunken",
            borderwidth=1,
            background=self.color_value.get(),
        )
        self.color_preview.grid(row=2, column=1, sticky="w", padx=(12, 8), pady=7)
        ttk.Button(frame, text="Pick Color...", command=self._pick_color).grid(
            row=2, column=2, sticky="w", pady=7
        )

        ttk.Label(frame, text="Link (optional)").grid(
            row=3, column=0, sticky="w", pady=7
        )
        ttk.Entry(frame, textvariable=self.link_value, width=44).grid(
            row=3, column=1, sticky="ew", padx=(12, 8), pady=7
        )
        ttk.Button(frame, text="Browse File...", command=self._browse_file).grid(
            row=3, column=2, sticky="w", pady=7
        )
        ttk.Label(
            frame,
            text="HTTPS links must match Allowed Domains; local executable files are blocked.",
            foreground="#555555",
        ).grid(row=4, column=1, columnspan=2, sticky="w", padx=(12, 0))

        buttons = ttk.Frame(frame)
        buttons.grid(row=5, column=0, columnspan=3, sticky="e", pady=(18, 0))
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(
            side="right", padx=(8, 0)
        )
        ttk.Button(buttons, text="Save Brick", command=self._save).pack(side="right")

        self.bind("<Return>", lambda _event: self._save())
        self.bind("<Escape>", lambda _event: self.destroy())
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.grab_set()
        name_entry.focus_set()
        self.update_idletasks()
        self._center_over_parent(parent.winfo_toplevel())

    def _pick_color(self) -> None:
        _, selected = colorchooser.askcolor(
            color=self.color_value.get(),
            parent=self,
            title="Choose brick color",
        )
        if selected:
            self.color_value.set(selected.upper())
            self.color_preview.configure(background=selected)

    def _browse_file(self) -> None:
        selected = filedialog.askopenfilename(parent=self, title="Select linked file")
        if selected:
            self.link_value.set(str(Path(selected).resolve()))

    def _save(self) -> None:
        name = self.name_value.get().strip()
        if not name:
            messagebox.showerror("Missing name", "Enter a name for the brick.", parent=self)
            return
        try:
            link = validate_link(self.link_value.get(), self.allowed_domains)
        except ValueError as error:
            messagebox.showerror("Invalid link", str(error), parent=self)
            return

        display_to_key = {label: key for key, label in COLUMN_LABELS.items()}
        self.result = {
            "name": name,
            "column": display_to_key[self.column_value.get()],
            "color": self.color_value.get().upper(),
            "link": link,
        }
        self.destroy()

    def _center_over_parent(self, parent) -> None:
        x = parent.winfo_rootx() + max(0, (parent.winfo_width() - self.winfo_width()) // 2)
        y = parent.winfo_rooty() + max(0, (parent.winfo_height() - self.winfo_height()) // 2)
        self.geometry(f"+{x}+{y}")


class AllowedDomainsDialog(tk.Toplevel):
    """Modal editor for the HTTPS domain allowlist."""

    def __init__(self, parent, allowed_domains: list[str]):
        super().__init__(parent)
        self.title("Allowed Web Domains")
        self.geometry("520x390")
        self.minsize(460, 330)
        self.transient(parent.winfo_toplevel())
        self.result: list[str] | None = None

        frame = ttk.Frame(self, padding=18)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        ttk.Label(
            frame,
            text="HTTPS Domain Allowlist",
            style="Heading.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            frame,
            text=(
                "Enter one domain per line. Use an exact domain such as "
                "example.com or an explicit wildcard such as *.example.com. "
                "Google's default entry accepts Sheets links only."
            ),
            wraplength=470,
        ).grid(row=1, column=0, sticky="ew", pady=(5, 10))

        text_frame = ttk.Frame(frame)
        text_frame.grid(row=2, column=0, sticky="nsew")
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        self.domain_text = tk.Text(text_frame, height=10, wrap="none")
        scrollbar = ttk.Scrollbar(
            text_frame,
            orient="vertical",
            command=self.domain_text.yview,
        )
        self.domain_text.configure(yscrollcommand=scrollbar.set)
        self.domain_text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.domain_text.insert("1.0", "\n".join(allowed_domains))

        buttons = ttk.Frame(frame)
        buttons.grid(row=3, column=0, sticky="e", pady=(14, 0))
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(
            side="right", padx=(8, 0)
        )
        ttk.Button(buttons, text="Save Domains", command=self._save).pack(
            side="right"
        )

        self.bind("<Escape>", lambda _event: self.destroy())
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.grab_set()
        self.domain_text.focus_set()

    def _save(self) -> None:
        values = [
            value.strip()
            for value in self.domain_text.get("1.0", "end").splitlines()
            if value.strip()
        ]
        try:
            self.result = normalize_allowed_domains(values)
        except ValueError as error:
            messagebox.showerror("Invalid domain", str(error), parent=self)
            return
        self.destroy()


class QueueBoard(ttk.Frame):
    """Three-column, draggable queue board."""

    CARD_HEIGHT = 68
    CARD_GAP = 9
    CARD_PADDING = 10

    def __init__(
        self,
        parent,
        data_file: str | Path,
        settings_file: str | Path,
        startup_command: str,
        status_callback=None,
    ):
        super().__init__(parent)
        self.status_callback = status_callback
        self.startup_command = startup_command
        self.canvases: dict[str, tk.Canvas] = {}
        self.drag_state: dict[str, object] | None = None
        self.layout_animation_id: str | None = None
        self.drop_animation_id: str | None = None
        self.animating_drop = False
        self.selected_brick: tuple[str, str] | None = None
        self.showing_archive = False
        self.storage_available = True
        self.settings_available = True

        self.settings = AppSettingsStore(settings_file)
        try:
            self.settings.load()
        except ValueError as error:
            self.settings_available = False
            self.after_idle(
                lambda message=str(error): messagebox.showerror(
                    "Security settings error",
                    f"{message}\n\nThe existing file was not overwritten.",
                    parent=self.winfo_toplevel(),
                )
            )

        self.store = QueueStore(data_file, self.settings.allowed_domains)

        try:
            self.store.load()
            if self.store.archived_during_load:
                self.store.save()
        except ValueError as error:
            self.storage_available = False
            self.after_idle(
                lambda message=str(error): messagebox.showerror(
                    "Queue data error",
                    f"{message}\n\nThe existing file was not overwritten.",
                    parent=self.winfo_toplevel(),
                )
            )
        except OSError as error:
            self.storage_available = False
            self.after_idle(
                lambda message=str(error): messagebox.showerror(
                    "Queue archive error",
                    f"Could not save the archived links: {message}",
                    parent=self.winfo_toplevel(),
                )
            )

        self.startup_value = tk.BooleanVar(
            value=get_startup_enabled(self.startup_command)
        )

        self._build_ui()
        self.after_idle(self._render_all)

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        toolbar = ttk.Frame(self)
        toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(toolbar, text="Queue Board", style="Heading.TLabel").pack(side="left")
        self.add_button = ttk.Button(
            toolbar,
            text="+ Add Brick",
            command=self._add_brick,
            state="normal" if self.storage_available else "disabled",
        )
        self.add_button.pack(side="right")
        self.archive_button = ttk.Button(
            toolbar,
            text="Show Archive",
            command=self._toggle_archive_view,
            state="normal" if self.storage_available else "disabled",
        )
        self.archive_button.pack(side="right", padx=(0, 8))
        self.domains_button = ttk.Button(
            toolbar,
            text="Allowed Domains...",
            command=self._edit_allowed_domains,
            state="normal" if self.settings_available else "disabled",
        )
        self.domains_button.pack(side="right", padx=(0, 8))
        ttk.Checkbutton(
            toolbar,
            text="Launch at sign-in",
            variable=self.startup_value,
            command=self._toggle_startup,
        ).pack(side="right", padx=(0, 12))

        self.board_instructions = tk.StringVar(
            value=(
                "Drag a brick to reorder its stack. Click a linked brick to open it. "
                "Right-click to edit, archive, or delete."
            )
        )
        ttk.Label(
            self,
            textvariable=self.board_instructions,
        ).grid(row=1, column=0, sticky="w", pady=(0, 10))

        board = ttk.Frame(self)
        board.grid(row=2, column=0, sticky="nsew")
        for index, column in enumerate(QUEUE_COLUMNS):
            board.columnconfigure(index, weight=1, uniform="queue")
        board.rowconfigure(0, weight=1)

        for index, column in enumerate(QUEUE_COLUMNS):
            panel = ttk.LabelFrame(board, text=COLUMN_LABELS[column], padding=5)
            panel.grid(
                row=0,
                column=index,
                sticky="nsew",
                padx=(0 if index == 0 else 5, 0 if index == 2 else 5),
            )
            panel.columnconfigure(0, weight=1)
            panel.rowconfigure(0, weight=1)

            canvas = tk.Canvas(
                panel,
                background="#D9D7CF",
                highlightthickness=0,
                borderwidth=0,
            )
            scrollbar = ttk.Scrollbar(panel, orient="vertical", command=canvas.yview)
            canvas.configure(yscrollcommand=scrollbar.set)
            canvas.grid(row=0, column=0, sticky="nsew")
            scrollbar.grid(row=0, column=1, sticky="ns")
            canvas.bind(
                "<Configure>", lambda _event, key=column: self._render_column(key)
            )
            canvas.bind(
                "<MouseWheel>",
                lambda event, target=canvas: target.yview_scroll(
                    -1 if event.delta > 0 else 1, "units"
                ),
            )
            self.canvases[column] = canvas

        self.context_menu = tk.Menu(self, tearoff=False)
        self.context_menu.add_command(label="Open Link", command=self._open_selected)
        self.context_menu.add_command(label="Edit...", command=self._edit_selected)
        self.context_menu.add_command(label="Archive", command=self._archive_selected)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Delete", command=self._delete_selected)

        self.archive_menu = tk.Menu(self, tearoff=False)
        self.archive_menu.add_command(label="Restore", command=self._restore_selected)
        self.archive_menu.add_command(label="Edit...", command=self._edit_selected)
        self.archive_menu.add_separator()
        self.archive_menu.add_command(label="Delete", command=self._delete_selected)

    def _render_all(self) -> None:
        for column in QUEUE_COLUMNS:
            self._render_column(column)
        archive_count = len(self.store.columns[ARCHIVE_COLUMN])
        if self.showing_archive:
            self.archive_button.configure(text=f"Back to Active ({archive_count})")
        else:
            self.archive_button.configure(text=f"Show Archive ({archive_count})")

    def _toggle_archive_view(self) -> None:
        if self.drag_state is not None or self.animating_drop:
            return
        self.showing_archive = not self.showing_archive
        self.drag_state = None
        self.selected_brick = None
        if self.showing_archive:
            self.add_button.configure(state="disabled")
            self.board_instructions.set(
                "Archive view: cards cannot be opened or dragged. "
                "Right-click to restore, edit, or delete."
            )
            self._set_status("Showing archived queue bricks.")
        else:
            self.add_button.configure(
                state="normal" if self.storage_available else "disabled"
            )
            self.board_instructions.set(
                "Drag a brick to reorder its stack. Click a linked brick to open it. "
                "Right-click to edit, archive, or delete."
            )
            self._set_status("Showing active queue bricks.")
        self._render_all()

    def _render_column(self, column: str) -> None:
        canvas = self.canvases.get(column)
        if canvas is None:
            return
        canvas.delete("all")
        width = max(canvas.winfo_width(), 220)
        card_width = max(120, width - self.CARD_PADDING * 2)
        if self.showing_archive:
            bricks = [
                brick
                for brick in self.store.columns[ARCHIVE_COLUMN]
                if brick.get("source_column", "general") == column
            ]
        else:
            bricks = self.store.columns[column]

        for index, brick in enumerate(bricks):
            top = self.CARD_PADDING + index * (self.CARD_HEIGHT + self.CARD_GAP)
            bottom = top + self.CARD_HEIGHT
            tag = f"brick_{brick['id']}"
            text_color = self._contrast_text(brick["color"])

            canvas.create_rectangle(
                self.CARD_PADDING + 3,
                top + 4,
                self.CARD_PADDING + card_width + 3,
                bottom + 4,
                fill="#777777",
                outline="",
                tags=(tag, "brick"),
            )
            canvas.create_rectangle(
                self.CARD_PADDING,
                top,
                self.CARD_PADDING + card_width,
                bottom,
                fill=brick["color"],
                outline="#484848",
                width=1,
                tags=(tag, "brick"),
            )
            if self.showing_archive:
                canvas.create_rectangle(
                    self.CARD_PADDING,
                    top,
                    self.CARD_PADDING + card_width,
                    bottom,
                    fill="#6E6E6E",
                    stipple="gray50",
                    outline="#484848",
                    width=1,
                    tags=(tag, "brick"),
                )
            canvas.create_text(
                self.CARD_PADDING + 12,
                top + 12,
                text=brick["name"],
                anchor="nw",
                width=max(80, card_width - 24),
                fill=text_color,
                font=("Segoe UI", 10, "bold"),
                tags=(tag, "brick"),
            )
            if self.showing_archive:
                archive_label = (
                    "LINK BLOCKED"
                    if brick.get("link_error") != "Archived manually."
                    else "ARCHIVED"
                )
                canvas.create_text(
                    self.CARD_PADDING + card_width - 10,
                    bottom - 10,
                    text=archive_label,
                    anchor="se",
                    fill=text_color,
                    font=("Segoe UI", 7, "bold"),
                    tags=(tag, "brick"),
                )
            elif brick["link"]:
                canvas.create_text(
                    self.CARD_PADDING + card_width - 10,
                    bottom - 10,
                    text="OPEN LINK",
                    anchor="se",
                    fill=text_color,
                    font=("Segoe UI", 7, "underline"),
                    tags=(tag, "brick"),
                )

            if self.showing_archive:
                canvas.tag_bind(
                    tag,
                    "<Button-3>",
                    lambda event, brick_id=brick["id"]: self._show_archive_context(
                        event, brick_id
                    ),
                )
            else:
                canvas.tag_bind(
                    tag,
                    "<ButtonPress-1>",
                    lambda event, key=column, brick_id=brick["id"]: self._drag_start(
                        event, key, brick_id
                    ),
                )
                canvas.tag_bind(tag, "<B1-Motion>", self._drag_motion)
                canvas.tag_bind(tag, "<ButtonRelease-1>", self._drag_release)
                canvas.tag_bind(
                    tag,
                    "<Button-3>",
                    lambda event, key=column, brick_id=brick["id"]: self._show_context(
                        event, key, brick_id
                    ),
                )

        total_height = max(
            canvas.winfo_height(),
            self.CARD_PADDING * 2
            + len(bricks) * (self.CARD_HEIGHT + self.CARD_GAP),
        )
        canvas.configure(scrollregion=(0, 0, width, total_height))

    @staticmethod
    def _contrast_text(color: str) -> str:
        red, green, blue = (
            int(color[1:3], 16),
            int(color[3:5], 16),
            int(color[5:7], 16),
        )
        luminance = 0.299 * red + 0.587 * green + 0.114 * blue
        return "#111111" if luminance > 155 else "#FFFFFF"

    def _slot_top(self, index: int) -> float:
        return self.CARD_PADDING + index * (self.CARD_HEIGHT + self.CARD_GAP)

    def _target_index(
        self,
        column: str,
        card_center: float,
        source_column: str,
    ) -> int:
        count = len(self.store.columns[column])
        if column == source_column:
            count -= 1
        raw_index = round(
            (card_center - self.CARD_PADDING - self.CARD_HEIGHT / 2)
            / (self.CARD_HEIGHT + self.CARD_GAP)
        )
        return max(0, min(raw_index, count))

    def _column_for_pointer(self, x_root: int) -> str:
        return min(
            QUEUE_COLUMNS,
            key=lambda column: abs(
                x_root
                - (
                    self.canvases[column].winfo_rootx()
                    + self.canvases[column].winfo_width() / 2
                )
            ),
        )

    def _create_drag_ghost(
        self,
        canvas: tk.Canvas,
        brick: dict[str, str],
        top: float,
    ) -> None:
        canvas.delete("drag_ghost")
        width = max(canvas.winfo_width(), 220)
        card_width = max(120, width - self.CARD_PADDING * 2)
        bottom = top + self.CARD_HEIGHT
        text_color = self._contrast_text(brick["color"])
        canvas.create_rectangle(
            self.CARD_PADDING + 5,
            top + 7,
            self.CARD_PADDING + card_width + 5,
            bottom + 7,
            fill="#555555",
            outline="",
            tags=("drag_ghost",),
        )
        canvas.create_rectangle(
            self.CARD_PADDING,
            top,
            self.CARD_PADDING + card_width,
            bottom,
            fill=brick["color"],
            outline="#222222",
            width=2,
            tags=("drag_ghost",),
        )
        canvas.create_text(
            self.CARD_PADDING + 12,
            top + 12,
            text=brick["name"],
            anchor="nw",
            width=max(80, card_width - 24),
            fill=text_color,
            font=("Segoe UI", 10, "bold"),
            tags=("drag_ghost",),
        )
        if brick["link"]:
            canvas.create_text(
                self.CARD_PADDING + card_width - 10,
                bottom - 10,
                text="MOVING",
                anchor="se",
                fill=text_color,
                font=("Segoe UI", 7, "bold"),
                tags=("drag_ghost",),
            )
        canvas.tag_raise("drag_ghost")

    def _create_drop_placeholder(self, column: str, index: int) -> None:
        for canvas in self.canvases.values():
            canvas.delete("drag_placeholder")
        canvas = self.canvases[column]
        width = max(canvas.winfo_width(), 220)
        card_width = max(120, width - self.CARD_PADDING * 2)
        top = self._slot_top(index)
        canvas.create_rectangle(
            self.CARD_PADDING,
            top,
            self.CARD_PADDING + card_width,
            top + self.CARD_HEIGHT,
            fill="#858585",
            stipple="gray50",
            outline="#555555",
            dash=(5, 3),
            width=2,
            tags=("drag_placeholder",),
        )
        canvas.create_text(
            self.CARD_PADDING + card_width / 2,
            top + self.CARD_HEIGHT / 2,
            text="DROP HERE",
            fill="#333333",
            font=("Segoe UI", 8, "bold"),
            tags=("drag_placeholder",),
        )

    def _cancel_layout_animation(self) -> None:
        if self.layout_animation_id is None:
            return
        try:
            self.after_cancel(self.layout_animation_id)
        except tk.TclError:
            pass
        self.layout_animation_id = None

    def _animate_drag_reflow(
        self,
        source_column: str,
        target_column: str,
        dragged_id: str,
        target_index: int,
    ) -> None:
        self._cancel_layout_animation()
        self._create_drop_placeholder(target_column, target_index)
        targets: list[tuple[tk.Canvas, str, float]] = []
        for column in QUEUE_COLUMNS:
            identifiers = [brick["id"] for brick in self.store.columns[column]]
            if column == source_column:
                identifiers = [
                    identifier for identifier in identifiers if identifier != dragged_id
                ]
            prospective: list[str | None] = list(identifiers)
            if column == target_column:
                prospective.insert(target_index, None)
            canvas = self.canvases[column]
            for index, identifier in enumerate(prospective):
                tag = (
                    "drag_placeholder"
                    if identifier is None
                    else f"brick_{identifier}"
                )
                targets.append((canvas, tag, self._slot_top(index)))

        starts: list[tuple[tk.Canvas, str, float, float]] = []
        for canvas, tag, target_top in targets:
            bounds = canvas.bbox(tag)
            if bounds is not None:
                starts.append((canvas, tag, float(bounds[1]), target_top))
        frame_count = 7

        def tick(frame: int = 1) -> None:
            progress = min(1.0, frame / frame_count)
            eased = 1 - (1 - progress) ** 3
            for canvas, tag, start_top, target_top in starts:
                bounds = canvas.bbox(tag)
                if bounds is None:
                    continue
                desired_top = start_top + (target_top - start_top) * eased
                canvas.move(tag, 0, desired_top - bounds[1])
            self.canvases[target_column].tag_raise("drag_ghost")
            if frame < frame_count:
                self.layout_animation_id = self.after(
                    16,
                    lambda: tick(frame + 1),
                )
            else:
                self.layout_animation_id = None

        tick()

    def _animate_drop_to_order(
        self,
        source_column: str,
        target_column: str,
        dragged_id: str,
        *,
        on_complete=None,
        duration_ms: int = 190,
    ) -> None:
        self._cancel_layout_animation()
        self.animating_drop = True
        targets: list[tuple[tk.Canvas, str, float]] = []
        for column in QUEUE_COLUMNS:
            for index, brick in enumerate(self.store.columns[column]):
                if brick["id"] == dragged_id:
                    item_canvas = self.canvases[target_column]
                    tag = "drag_ghost"
                else:
                    item_canvas = self.canvases[column]
                    tag = f"brick_{brick['id']}"
                targets.append((item_canvas, tag, self._slot_top(index)))

        starts: list[tuple[tk.Canvas, str, float, float]] = []
        for canvas, tag, target_top in targets:
            bounds = canvas.bbox(tag)
            if bounds is not None:
                starts.append((canvas, tag, float(bounds[1]), target_top))
        frame_count = max(1, round(duration_ms / 16))

        def finish() -> None:
            self.drop_animation_id = None
            self.animating_drop = False
            for canvas in self.canvases.values():
                canvas.configure(cursor="")
                canvas.delete("drag_placeholder")
                canvas.delete("drag_ghost")
            self._render_all()
            if on_complete is not None:
                on_complete()

        def tick(frame: int = 1) -> None:
            progress = min(1.0, frame / frame_count)
            eased = 1 - (1 - progress) ** 3
            for canvas, tag, start_top, target_top in starts:
                bounds = canvas.bbox(tag)
                if bounds is None:
                    continue
                desired_top = start_top + (target_top - start_top) * eased
                canvas.move(tag, 0, desired_top - bounds[1])
            self.canvases[target_column].tag_raise("drag_ghost")
            if frame < frame_count:
                self.drop_animation_id = self.after(
                    16,
                    lambda: tick(frame + 1),
                )
            else:
                finish()

        tick()

    def _drag_start(self, event, column: str, brick_id: str) -> None:
        if self.showing_archive or self.animating_drop:
            return
        canvas = self.canvases[column]
        found = self.store.find(brick_id)
        if found is None or found[0] != column:
            return
        self._cancel_layout_animation()
        start_y = canvas.canvasy(event.y)
        original_index = found[1]
        placeholder_top = self._slot_top(original_index)
        self._create_drop_placeholder(column, original_index)
        self._create_drag_ghost(canvas, found[2], placeholder_top)
        canvas.itemconfigure(f"brick_{brick_id}", state="hidden")
        self.drag_state = {
            "source_column": column,
            "target_column": column,
            "ghost_column": column,
            "brick_id": brick_id,
            "start_root_x": event.x_root,
            "start_root_y": event.y_root,
            "grab_offset": start_y - placeholder_top,
            "moved": False,
            "original_index": original_index,
            "target_index": original_index,
        }
        canvas.configure(cursor="fleur")
        canvas.tag_raise("drag_ghost")

    def _drag_motion(self, event) -> None:
        if self.drag_state is None:
            return
        source_column = str(self.drag_state["source_column"])
        target_column = self._column_for_pointer(event.x_root)
        target_canvas = self.canvases[target_column]
        local_y = event.y_root - target_canvas.winfo_rooty()
        pointer_y = target_canvas.canvasy(local_y)
        desired_top = pointer_y - float(self.drag_state["grab_offset"])
        if (
            abs(event.x_root - int(self.drag_state["start_root_x"])) >= 4
            or abs(event.y_root - int(self.drag_state["start_root_y"])) >= 4
        ):
            self.drag_state["moved"] = True

        brick_id = str(self.drag_state["brick_id"])
        ghost_column = str(self.drag_state["ghost_column"])
        if ghost_column != target_column:
            self.canvases[ghost_column].delete("drag_ghost")
            self.canvases[ghost_column].configure(cursor="")
            found = self.store.find(brick_id)
            if found is None:
                return
            self._create_drag_ghost(target_canvas, found[2], desired_top)
            target_canvas.configure(cursor="fleur")
            self.drag_state["ghost_column"] = target_column
        else:
            bounds = target_canvas.bbox("drag_ghost")
            if bounds is not None:
                target_canvas.move("drag_ghost", 0, desired_top - bounds[1])

        target_index = self._target_index(
            target_column,
            desired_top + self.CARD_HEIGHT / 2,
            source_column,
        )
        if (
            target_column != str(self.drag_state["target_column"])
            or target_index != int(self.drag_state["target_index"])
        ):
            self.drag_state["target_column"] = target_column
            self.drag_state["target_index"] = target_index
            self._animate_drag_reflow(
                source_column,
                target_column,
                brick_id,
                target_index,
            )
        target_canvas.tag_raise("drag_ghost")

        if local_y < 18:
            target_canvas.yview_scroll(-1, "units")
        elif local_y > target_canvas.winfo_height() - 18:
            target_canvas.yview_scroll(1, "units")

    def _drag_release(self, event) -> None:
        if self.drag_state is None:
            return
        state = self.drag_state
        self.drag_state = None
        source_column = str(state["source_column"])
        target_column = str(state["target_column"])
        brick_id = str(state["brick_id"])

        if not bool(state["moved"]):
            self._animate_drop_to_order(
                source_column,
                source_column,
                brick_id,
                on_complete=lambda: self._open_brick(brick_id),
                duration_ms=110,
            )
            return

        target_index = int(state["target_index"])
        snapshot = self.store.snapshot()
        try:
            changed = self.store.move(brick_id, target_column, target_index)
            if changed:
                self.store.save()
        except (OSError, KeyError, ValueError) as error:
            self.store.restore(snapshot)
            self._cancel_layout_animation()
            for canvas in self.canvases.values():
                canvas.configure(cursor="")
                canvas.delete("drag_placeholder")
                canvas.delete("drag_ghost")
            messagebox.showerror("Could not save", str(error), parent=self.winfo_toplevel())
            self._render_all()
            return
        status_message = (
            f"Brick moved to {COLUMN_LABELS[target_column]}."
            if source_column != target_column
            else "Queue order saved."
        )
        self._animate_drop_to_order(
            source_column,
            target_column,
            brick_id,
            on_complete=(
                lambda: self._set_status(status_message)
                if changed
                else None
            ),
        )

    def _add_brick(self) -> None:
        dialog = BrickDialog(
            self,
            "Add Queue Brick",
            allowed_domains=self.settings.allowed_domains,
        )
        self.wait_window(dialog)
        if dialog.result is None:
            return

        snapshot = self.store.snapshot()
        try:
            self.store.add(**dialog.result)
            self.store.save()
        except (OSError, ValueError) as error:
            self.store.restore(snapshot)
            messagebox.showerror("Could not save", str(error), parent=self.winfo_toplevel())
            return
        self._render_all()
        self._set_status("Brick added and queue saved.")

    def _edit_allowed_domains(self) -> None:
        dialog = AllowedDomainsDialog(self, self.settings.allowed_domains)
        self.wait_window(dialog)
        if dialog.result is None:
            return

        snapshot = self.settings.snapshot()
        try:
            self.settings.allowed_domains = dialog.result
            self.settings.save()
        except OSError as error:
            self.settings.restore(snapshot)
            messagebox.showerror(
                "Could not save domains",
                str(error),
                parent=self.winfo_toplevel(),
            )
            return
        self.store.allowed_domains = list(self.settings.allowed_domains)
        self._set_status("Allowed web domains saved.")

    def _toggle_startup(self) -> None:
        requested = self.startup_value.get()
        try:
            set_startup_enabled(self.startup_command, requested)
        except OSError as error:
            self.startup_value.set(not requested)
            messagebox.showerror(
                "Could not update startup",
                str(error),
                parent=self.winfo_toplevel(),
            )
            return
        state = "enabled" if requested else "disabled"
        self._set_status(f"Launch at sign-in {state}.")

    def _show_context(self, event, column: str, brick_id: str) -> None:
        self.selected_brick = (column, brick_id)
        found = self.store.find(brick_id)
        link_state = "normal" if found and found[2]["link"] else "disabled"
        self.context_menu.entryconfigure("Open Link", state=link_state)
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def _show_archive_context(self, event, brick_id: str) -> None:
        self.selected_brick = (ARCHIVE_COLUMN, brick_id)
        try:
            self.archive_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.archive_menu.grab_release()

    def _open_selected(self) -> None:
        if self.selected_brick is not None:
            self._open_brick(self.selected_brick[1])

    def _archive_selected(self) -> None:
        if self.selected_brick is None:
            return
        snapshot = self.store.snapshot()
        try:
            self.store.archive(self.selected_brick[1])
            self.store.save()
        except (OSError, KeyError) as error:
            self.store.restore(snapshot)
            messagebox.showerror(
                "Could not archive",
                str(error),
                parent=self.winfo_toplevel(),
            )
            return
        self._render_all()
        self._set_status("Brick archived.")

    def _restore_selected(self) -> None:
        if self.selected_brick is None:
            return
        snapshot = self.store.snapshot()
        try:
            target_column = self.store.unarchive(self.selected_brick[1])
            self.store.save()
        except (OSError, KeyError, ValueError) as error:
            self.store.restore(snapshot)
            messagebox.showerror(
                "Could not restore",
                f"{error}\n\nEdit the brick to repair or remove its link.",
                parent=self.winfo_toplevel(),
            )
            return
        self._render_all()
        self._set_status(f"Brick restored to {COLUMN_LABELS[target_column]}.")

    def _edit_selected(self) -> None:
        if self.selected_brick is None:
            return
        found = self.store.find(self.selected_brick[1])
        if found is None:
            return
        column, _, brick = found
        initial_column = (
            brick.get("source_column", "general")
            if column == ARCHIVE_COLUMN
            else column
        )
        dialog = BrickDialog(
            self,
            "Edit Queue Brick",
            allowed_domains=self.settings.allowed_domains,
            initial_column=initial_column,
            initial_brick=brick,
        )
        self.wait_window(dialog)
        if dialog.result is None:
            return

        snapshot = self.store.snapshot()
        try:
            self.store.update(brick["id"], **dialog.result)
            self.store.save()
        except (OSError, KeyError, ValueError) as error:
            self.store.restore(snapshot)
            messagebox.showerror("Could not save", str(error), parent=self.winfo_toplevel())
            return
        self._render_all()
        self._set_status("Brick updated and queue saved.")

    def _delete_selected(self) -> None:
        if self.selected_brick is None:
            return
        found = self.store.find(self.selected_brick[1])
        if found is None:
            return
        brick = found[2]
        confirmed = messagebox.askyesno(
            "Delete brick",
            f"Delete '{brick['name']}' from the queue?",
            parent=self.winfo_toplevel(),
        )
        if not confirmed:
            return

        snapshot = self.store.snapshot()
        try:
            self.store.delete(brick["id"])
            self.store.save()
        except (OSError, KeyError) as error:
            self.store.restore(snapshot)
            messagebox.showerror("Could not save", str(error), parent=self.winfo_toplevel())
            return
        self._render_all()
        self._set_status("Brick deleted and queue saved.")

    def _open_brick(self, brick_id: str) -> None:
        found = self.store.find(brick_id)
        if (
            found is None
            or found[0] == ARCHIVE_COLUMN
            or not found[2]["link"]
        ):
            return
        link = found[2]["link"]
        path = Path(link)
        try:
            if path.is_absolute():
                validate_link(link, self.settings.allowed_domains)
                resolved = path.resolve(strict=True)
                if not resolved.is_file():
                    raise FileNotFoundError(f"Linked file was not found:\n{path}")
                validate_link(str(resolved), self.settings.allowed_domains)
                os.startfile(resolved)
            else:
                validate_link(link, self.settings.allowed_domains)
                opened = webbrowser.open(link, new=2)
                if not opened:
                    raise OSError("Windows could not open the URL.")
        except (OSError, ValueError) as error:
            messagebox.showerror("Could not open link", str(error), parent=self.winfo_toplevel())

    def _set_status(self, message: str) -> None:
        if self.status_callback is not None:
            self.status_callback(message)
