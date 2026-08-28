"""
Graphical User Interface for Entomology Labels Generator.

Provides an easy-to-use interface for creating and exporting entomology labels.
"""

import logging
import math
import queue
import tempfile
import threading
import tkinter as tk
import webbrowser
from collections import Counter
from pathlib import Path
from tkinter import filedialog
from tkinter import font as tkfont
from tkinter import messagebox, ttk
from typing import Optional

from .config import (
    FONT_SIZE_PT_MAX,
    FONT_SIZE_PT_MIN,
    LABEL_HEIGHT_MM_MAX,
    LABEL_HEIGHT_MM_MIN,
    LABEL_WIDTH_MM_MAX,
    LABEL_WIDTH_MM_MIN,
    MARGIN_MM_MAX,
    MARGIN_MM_MIN,
    MAX_DISPLAYED_LABELS,
    MAX_LABELS_PER_GENERATOR,
    PREVIEW_SCALE_FACTOR,
)
from .fit import MM_PER_POINT
from .input_handlers import load_data
from .label_generator import Label, LabelConfig, LabelGenerator, expand_label
from .layout import STYLE_SPACER, render_label_lines
from .output_generators import generate_docx, generate_html, generate_pdf

logger = logging.getLogger(__name__)

# How often the main loop checks whether a background job has finished
BACKGROUND_POLL_MS = 50


def _bind_mousewheel(widget, canvas) -> None:
    """Scroll `canvas` when the wheel is used over `widget`.

    Bound per widget rather than with bind_all so each scrollable area only
    responds to the wheel while the pointer is actually over it.
    """

    def _on_wheel(event):
        if event.num == 4:  # X11 wheel up
            delta = -1
        elif event.num == 5:  # X11 wheel down
            delta = 1
        else:
            delta = int(-1 * (event.delta / 120))
        canvas.yview_scroll(delta, "units")
        return "break"

    for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
        widget.bind(sequence, _on_wheel)


# A preview line should occupy only the space its text occupies on paper. A
# tk.Label adds a border and internal padding by default, which is about a
# third of a line at preview scale, so it is turned off.
_PREVIEW_LINE_KW = {"bd": 0, "padx": 0, "pady": 0, "highlightthickness": 0}


def _preview_font_px(config: LabelConfig, scale: float, measure=None) -> int:
    """Pick a preview font size whose lines match the printed line height.

    tkinter reads a negative font size as a height in pixels, which is what
    makes the text commensurate with a box measured in scaled millimetres.
    A font's rendered line height exceeds its nominal size by an amount that
    varies with the face, so rather than guess a factor, ask the font for its
    own metrics and step down until a line fits.

    Args:
        config: Layout configuration
        scale: Preview pixels per millimetre
        measure: Returns the rendered line height for a pixel size. Defaults
            to querying tkinter; injectable so the sizing logic can be tested
            without a display.

    Returns:
        A positive pixel size; pass it to tkinter negated
    """
    target_px = config.font_size_pt * config.line_spacing * MM_PER_POINT * scale
    px = max(1, int(round(config.font_size_pt * MM_PER_POINT * scale)))

    if measure is None:

        def measure(size: int) -> int:
            return tkfont.Font(family=config.font_family, size=-size).metrics("linespace")

    while px > 1:
        try:
            linespace = measure(px)
        except tk.TclError:  # pragma: no cover - no font server available
            break
        if linespace <= target_px:
            break
        px -= 1

    return px


class EntomologyLabelsGUI:
    """Main GUI application for generating entomology labels."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Entomology Labels Generator")
        self.root.geometry("1100x800")
        self.root.minsize(900, 700)

        # Initialize generator
        self.generator = LabelGenerator()

        # Index of the label currently loaded in the form for editing, if any
        self._editing_index: Optional[int] = None

        # Current page of the label list
        self._tree_page = 0

        # Set while a background import or export is running
        self._busy = False

        # Setup UI
        self._setup_menu()
        self._setup_main_layout()
        self._setup_bindings()

        # Center window
        self._center_window()

    def _center_window(self):
        """Center the window on screen."""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _setup_menu(self):
        """Setup the menu bar."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(
            label="Import Data...", command=self._import_data, accelerator="Ctrl+O"
        )
        file_menu.add_separator()
        file_menu.add_command(label="Export HTML...", command=lambda: self._export("html"))
        file_menu.add_command(label="Export PDF...", command=lambda: self._export("pdf"))
        file_menu.add_command(label="Export DOCX...", command=lambda: self._export("docx"))
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit, accelerator="Ctrl+Q")

        # Edit menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Clear All Labels", command=self._clear_labels)
        edit_menu.add_command(
            label="Generate Sequential Labels...", command=self._show_sequential_dialog
        )

        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Guide", command=self._show_help)
        help_menu.add_command(label="About", command=self._show_about)

    def _setup_main_layout(self):
        """Setup the main application layout."""
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Tab 1: Label Data
        self._setup_data_tab()

        # Tab 2: Visual Preview
        self._setup_preview_tab()

        # Tab 3: Configuration
        self._setup_config_tab()

        # Bottom status bar
        self._setup_status_bar(main_frame)

    def _setup_data_tab(self):
        """Setup the data entry tab."""
        data_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(data_frame, text="Label Data")

        # Top panel - Import and Quick Actions
        top_frame = ttk.LabelFrame(data_frame, text="Import & Actions", padding="10")
        top_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))

        ttk.Button(top_frame, text="Import from File...", command=self._import_data).pack(
            side=tk.LEFT, padx=5
        )
        ttk.Button(
            top_frame, text="Generate Sequential...", command=self._show_sequential_dialog
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="Clear All", command=self._clear_labels).pack(
            side=tk.RIGHT, padx=5
        )

        # PanedWindow for entry form and list
        paned = ttk.PanedWindow(data_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # Left panel - Form for single label entry
        left_frame = ttk.LabelFrame(paned, text="Add Single Label", padding="10")
        paned.add(left_frame, weight=1)

        # Form fields
        fields = [
            ("Location (Line 1):", "location1"),
            ("Location (Line 2):", "location2"),
            ("Code:", "code"),
            ("Date:", "date"),
            ("Additional Notes:", "notes"),
            ("Quantity:", "quantity"),
        ]

        self.entry_vars = {}
        for i, (label_text, var_name) in enumerate(fields):
            ttk.Label(left_frame, text=label_text).grid(row=i, column=0, sticky=tk.W, pady=5)
            var = tk.StringVar()
            self.entry_vars[var_name] = var
            if var_name == "quantity":
                var.set("1")
            entry = ttk.Entry(left_frame, textvariable=var)
            entry.grid(row=i, column=1, sticky=tk.EW, pady=5, padx=(5, 0))

        left_frame.columnconfigure(1, weight=1)

        # Buttons
        btn_frame = ttk.Frame(left_frame)
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=15)

        ttk.Button(btn_frame, text="Add Label", command=self._add_label).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Clear Form", command=self._clear_form).pack(
            side=tk.LEFT, padx=5
        )

        # Right panel - Labels list
        right_frame = ttk.LabelFrame(paned, text="Label List", padding="10")
        paned.add(right_frame, weight=2)

        # Treeview for labels
        columns = ("location1", "location2", "code", "date", "quantity")
        self.labels_tree = ttk.Treeview(right_frame, columns=columns, show="headings")

        self.labels_tree.heading("location1", text="Location 1")
        self.labels_tree.heading("location2", text="Location 2")
        self.labels_tree.heading("code", text="Code")
        self.labels_tree.heading("date", text="Date")
        self.labels_tree.heading("quantity", text="Copies")

        self.labels_tree.column("location1", width=150)
        self.labels_tree.column("location2", width=150)
        self.labels_tree.column("code", width=70)
        self.labels_tree.column("date", width=90)
        self.labels_tree.column("quantity", width=55)

        # Scrollbar
        scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.labels_tree.yview)
        self.labels_tree.configure(yscrollcommand=scrollbar.set)

        self.labels_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Double-click to edit
        self.labels_tree.bind("<Double-1>", lambda e: self._edit_selected_label())

        # Buttons under treeview
        tree_btn_frame = ttk.Frame(right_frame)
        tree_btn_frame.pack(fill=tk.X, pady=(5, 0))

        ttk.Button(
            tree_btn_frame, text="Remove Selected", command=self._remove_selected_label
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            tree_btn_frame, text="Duplicate Selected", command=self._duplicate_selected_label
        ).pack(side=tk.LEFT, padx=2)

        # Paging controls, so every label stays reachable however many there are
        self.tree_next_button = ttk.Button(
            tree_btn_frame, text="Next >", command=lambda: self._change_tree_page(1)
        )
        self.tree_next_button.pack(side=tk.RIGHT, padx=2)
        self.tree_page_label = ttk.Label(tree_btn_frame, text="No labels")
        self.tree_page_label.pack(side=tk.RIGHT, padx=8)
        self.tree_prev_button = ttk.Button(
            tree_btn_frame, text="< Prev", command=lambda: self._change_tree_page(-1)
        )
        self.tree_prev_button.pack(side=tk.RIGHT, padx=2)

    def _setup_preview_tab(self):
        """Setup the visual preview tab."""
        preview_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(preview_frame, text="Visual Preview")

        # Top controls
        controls_frame = ttk.Frame(preview_frame)
        controls_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(controls_frame, text="Refresh Preview", command=self._update_preview).pack(
            side=tk.LEFT, padx=5
        )

        ttk.Label(controls_frame, text="Page:").pack(side=tk.LEFT, padx=(20, 5))
        self.page_var = tk.StringVar(value="1")
        self.page_spinbox = ttk.Spinbox(
            controls_frame,
            from_=1,
            to=1,
            textvariable=self.page_var,
            width=5,
            command=self._update_preview,
        )
        self.page_spinbox.pack(side=tk.LEFT)
        self.total_pages_label = ttk.Label(controls_frame, text="of 0")
        self.total_pages_label.pack(side=tk.LEFT, padx=5)

        ttk.Separator(controls_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=15)

        ttk.Button(controls_frame, text="Export PDF", command=lambda: self._export("pdf")).pack(
            side=tk.RIGHT, padx=5
        )
        ttk.Button(controls_frame, text="Export HTML", command=lambda: self._export("html")).pack(
            side=tk.RIGHT, padx=5
        )
        ttk.Button(controls_frame, text="Export DOCX", command=lambda: self._export("docx")).pack(
            side=tk.RIGHT, padx=5
        )

        # Canvas for preview
        canvas_frame = ttk.Frame(preview_frame, relief=tk.SUNKEN, borderwidth=1)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.preview_canvas = tk.Canvas(canvas_frame, bg="gray")
        self.preview_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        v_scroll = ttk.Scrollbar(
            canvas_frame, orient=tk.VERTICAL, command=self.preview_canvas.yview
        )
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll = ttk.Scrollbar(
            preview_frame, orient=tk.HORIZONTAL, command=self.preview_canvas.xview
        )
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        self.preview_canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        # Inner frame for the "paper"
        self.paper_frame = tk.Frame(self.preview_canvas, bg="white")
        self.preview_canvas.create_window((10, 10), window=self.paper_frame, anchor="nw")

        _bind_mousewheel(self.preview_canvas, self.preview_canvas)
        _bind_mousewheel(self.paper_frame, self.preview_canvas)

    def _setup_config_tab(self):
        """Setup the configuration tab."""
        container = ttk.Frame(self.notebook)
        self.notebook.add(container, text="Configuration")

        config_canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=config_canvas.yview)
        scrollable_frame = ttk.Frame(config_canvas, padding="20")

        scrollable_frame.bind(
            "<Configure>", lambda e: config_canvas.configure(scrollregion=config_canvas.bbox("all"))
        )

        config_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        config_canvas.configure(yscrollcommand=scrollbar.set)

        config_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Wheel scrolling is bound to the canvas and its children rather than
        # with bind_all, which would hijack the wheel for the whole app and
        # leave the preview canvas unscrollable.
        _bind_mousewheel(config_canvas, config_canvas)
        _bind_mousewheel(scrollable_frame, config_canvas)

        # Grouping fields
        # Layout
        layout_group = ttk.LabelFrame(scrollable_frame, text="Label Layout", padding="15")
        layout_group.pack(fill=tk.X, pady=10)

        layout_fields = [
            ("Labels per row:", "labels_per_row", "10"),
            ("Labels per column:", "labels_per_column", "13"),
            ("Label width (mm):", "label_width_mm", "29.0"),
            ("Label height (mm):", "label_height_mm", "13.0"),
        ]

        self.config_vars = {}
        for i, (label_text, var_name, default) in enumerate(layout_fields):
            ttk.Label(layout_group, text=label_text).grid(row=i, column=0, sticky=tk.W, pady=5)
            var = tk.StringVar(value=default)
            self.config_vars[var_name] = var
            ttk.Entry(layout_group, textvariable=var, width=15).grid(
                row=i, column=1, sticky=tk.W, pady=5, padx=10
            )

        # Page
        page_group = ttk.LabelFrame(scrollable_frame, text="Page & Margins", padding="15")
        page_group.pack(fill=tk.X, pady=10)

        page_fields = [
            ("Page width (mm):", "page_width_mm", "297"),
            ("Page height (mm):", "page_height_mm", "210"),
            ("Top margin (mm):", "margin_top_mm", "0"),
            ("Bottom margin (mm):", "margin_bottom_mm", "0"),
            ("Left margin (mm):", "margin_left_mm", "0"),
            ("Right margin (mm):", "margin_right_mm", "0"),
        ]

        for i, (label_text, var_name, default) in enumerate(page_fields):
            ttk.Label(page_group, text=label_text).grid(row=i, column=0, sticky=tk.W, pady=5)
            var = tk.StringVar(value=default)
            self.config_vars[var_name] = var
            ttk.Entry(page_group, textvariable=var, width=15).grid(
                row=i, column=1, sticky=tk.W, pady=5, padx=10
            )

        # Font
        font_group = ttk.LabelFrame(scrollable_frame, text="Typography", padding="15")
        font_group.pack(fill=tk.X, pady=10)

        ttk.Label(font_group, text="Font Family:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.config_vars["font_family"] = tk.StringVar(value="Arial")
        font_combo = ttk.Combobox(
            font_group,
            textvariable=self.config_vars["font_family"],
            values=["Arial", "Times New Roman", "Helvetica", "Calibri", "Courier New"],
            width=20,
        )
        font_combo.grid(row=0, column=1, sticky=tk.W, pady=5, padx=10)

        ttk.Label(font_group, text="Font Size (pt):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.config_vars["font_size_pt"] = tk.StringVar(value="6")
        ttk.Entry(font_group, textvariable=self.config_vars["font_size_pt"], width=15).grid(
            row=1, column=1, sticky=tk.W, pady=5, padx=10
        )

        ttk.Label(font_group, text="Line Spacing:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.config_vars["line_spacing"] = tk.StringVar(value="1.0")
        ttk.Entry(font_group, textvariable=self.config_vars["line_spacing"], width=15).grid(
            row=2, column=1, sticky=tk.W, pady=5, padx=10
        )

        # Buttons
        actions_frame = ttk.Frame(scrollable_frame)
        actions_frame.pack(fill=tk.X, pady=20)

        ttk.Button(actions_frame, text="Apply Changes", command=self._apply_config).pack(
            side=tk.LEFT, padx=5
        )

        # Presets
        preset_group = ttk.LabelFrame(scrollable_frame, text="Presets", padding="15")
        preset_group.pack(fill=tk.X, pady=10)

        ttk.Button(
            preset_group,
            text="A4 Landscape (10x13)",
            command=lambda: self._apply_preset("a4_standard"),
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            preset_group,
            text="A4 Compact (12x15)",
            command=lambda: self._apply_preset("a4_compact"),
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            preset_group, text="US Letter (10x12)", command=lambda: self._apply_preset("letter_us")
        ).pack(side=tk.LEFT, padx=5)

    def _setup_status_bar(self, parent):
        """Setup the status bar."""
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill=tk.X, pady=(10, 0))

        self.status_label = ttk.Label(status_frame, text="Ready")
        self.status_label.pack(side=tk.LEFT)

        self.labels_count_label = ttk.Label(status_frame, text="Labels: 0 | Pages: 0")
        self.labels_count_label.pack(side=tk.RIGHT)

    def _setup_bindings(self):
        """Setup keyboard bindings."""
        self.root.bind("<Control-o>", lambda e: self._import_data())
        self.root.bind("<Control-q>", lambda e: self.root.quit())
        self.root.bind("<Control-s>", lambda e: self._export("pdf"))

    def _add_label(self):
        """Add a label from the form."""
        try:
            qty_str = self.entry_vars["quantity"].get().strip()
            quantity = int(qty_str) if qty_str else 1
            if quantity <= 0:
                quantity = 1
        except ValueError:
            quantity = 1

        label = Label(
            location_line1=self.entry_vars["location1"].get(),
            location_line2=self.entry_vars["location2"].get(),
            code=self.entry_vars["code"].get(),
            date=self.entry_vars["date"].get(),
            additional_info=self.entry_vars["notes"].get(),
        )

        if label.is_empty():
            messagebox.showwarning("Warning", "Please fill at least one field for the label.")
            return

        copies = expand_label(label, quantity)

        editing_index = self._editing_index
        if editing_index is not None and editing_index < len(self.generator.labels):
            # Saving an edit: replace the original in place, keeping its position
            if len(self.generator.labels) - 1 + quantity > MAX_LABELS_PER_GENERATOR:
                messagebox.showerror(
                    "Error",
                    f"Maximum label count ({MAX_LABELS_PER_GENERATOR}) exceeded.",
                )
                return
            self.generator.labels[editing_index : editing_index + 1] = copies
            status = f"Updated label ({quantity} copy/copies)"
        else:
            try:
                for copy in copies:
                    self.generator.add_label(copy)
            except ValueError as e:
                messagebox.showerror("Error", str(e))
                return
            status = f"Added {quantity} label(s)"

        self._update_labels_tree()
        self._clear_form()
        self._update_status(status)

        self._refresh_preview_if_visible()

    def _clear_form(self):
        """Clear the entry form and abandon any in-progress edit."""
        self._editing_index = None
        for var_name, var in self.entry_vars.items():
            if var_name == "quantity":
                var.set("1")
            else:
                var.set("")

    def _run_in_background(self, work, on_success, on_error, status: str) -> None:
        """Run `work` off the Tk main loop and deliver the result back on it.

        Loading a large file and rendering a PDF both take long enough to
        freeze the window if they run inline. The worker hands its result over
        through a queue that the main thread polls, so no Tk call is ever made
        from another thread.

        Args:
            work: Callable executed on the worker thread
            on_success: Called on the main thread with the work's return value
            on_error: Called on the main thread with the raised exception
            status: Message shown in the status bar while running
        """
        if self._busy:
            messagebox.showinfo("Please wait", "Another operation is still running.")
            return

        self._busy = True
        self._update_status(status)
        self.root.config(cursor="watch")

        results: "queue.Queue" = queue.Queue(maxsize=1)

        def runner():
            try:
                results.put(("ok", work()))
            except Exception as exc:  # reported via on_error on the main thread
                results.put(("error", exc))

        threading.Thread(target=runner, daemon=True).start()
        self._poll_background(results, on_success, on_error)

    def _poll_background(self, results, on_success, on_error) -> None:
        """Check for a finished background job, rescheduling until it lands."""
        try:
            outcome, payload = results.get_nowait()
        except queue.Empty:
            self.root.after(
                BACKGROUND_POLL_MS,
                lambda: self._poll_background(results, on_success, on_error),
            )
            return

        self._busy = False
        self.root.config(cursor="")

        if outcome == "ok":
            on_success(payload)
        else:
            on_error(payload)

    def _refresh_preview_if_visible(self) -> None:
        """Re-render the preview when its tab is the one on screen."""
        if self.notebook.index(self.notebook.select()) == 1:
            self._update_preview()

    def _import_data(self):
        """Import data from a file."""
        filetypes = [
            ("All Supported Formats", "*.xlsx *.xls *.csv *.txt *.docx *.json *.yaml *.yml"),
            ("Excel", "*.xlsx *.xls"),
            ("CSV", "*.csv"),
            ("Text", "*.txt"),
            ("Word", "*.docx"),
            ("JSON", "*.json"),
            ("YAML", "*.yaml *.yml"),
        ]

        file_path = filedialog.askopenfilename(title="Select File to Import", filetypes=filetypes)

        if not file_path:
            return

        logger.info(f"Importing data from: {file_path}")

        def on_success(labels):
            # add_labels enforces the maximum, so this can still raise
            try:
                self.generator.add_labels(labels)
            except ValueError as e:
                messagebox.showerror("Import Error", f"Failed to import data:\n{e}")
                return
            self._update_labels_tree()
            logger.info(f"Successfully imported {len(labels)} labels")
            self._update_status(f"Imported {len(labels)} labels from {Path(file_path).name}")
            self.notebook.select(0)
            self._refresh_preview_if_visible()

        def on_error(exc):
            logger.error(f"Failed to import data: {exc}", exc_info=exc)
            messagebox.showerror("Import Error", f"Failed to import data:\n{exc}")

        self._run_in_background(
            lambda: load_data(file_path),
            on_success,
            on_error,
            status=f"Importing {Path(file_path).name}...",
        )

    @property
    def _tree_page_count(self) -> int:
        """Number of pages the label list is split into."""
        total = self.generator.total_labels
        return max(1, math.ceil(total / MAX_DISPLAYED_LABELS))

    def _change_tree_page(self, delta: int) -> None:
        """Move the label list one page forward or back."""
        new_page = self._tree_page + delta
        if 0 <= new_page < self._tree_page_count:
            self._tree_page = new_page
            self._update_labels_tree()

    @staticmethod
    def _label_key(label) -> tuple:
        """Return a hashable identity for a label, for counting duplicates."""
        return tuple(sorted(label.to_dict().items()))

    def _update_labels_tree(self):
        """Update the labels treeview.

        Only one page of MAX_DISPLAYED_LABELS rows is inserted at a time, which
        keeps the widget responsive while still letting every label be selected
        and edited — the previous version collapsed everything past the first
        500 into a single unselectable row.
        """
        # Clear existing items
        for item in self.labels_tree.get_children():
            self.labels_tree.delete(item)

        # A requested quantity is expanded into that many Label objects when
        # the label is added, so there is no per-row quantity to read back.
        # Counting identical labels recovers the same information, and unlike
        # the hardcoded "1" this column used to show, it is true.
        copies = Counter(self._label_key(label) for label in self.generator.labels)

        # Clamp the page in case labels were removed since the last refresh
        self._tree_page = max(0, min(self._tree_page, self._tree_page_count - 1))

        start = self._tree_page * MAX_DISPLAYED_LABELS
        end = min(start + MAX_DISPLAYED_LABELS, self.generator.total_labels)

        # The row id is the label's index in the generator, so selection-based
        # actions address the right label on any page.
        for i in range(start, end):
            label = self.generator.labels[i]
            self.labels_tree.insert(
                "",
                tk.END,
                iid=str(i),
                values=(
                    label.location_line1[:30] + ("..." if len(label.location_line1) > 30 else ""),
                    label.location_line2[:30] + ("..." if len(label.location_line2) > 30 else ""),
                    label.code,
                    label.date,
                    str(copies.get(self._label_key(label), 1)),
                ),
            )

        # Update paging controls
        page_count = self._tree_page_count
        if self.generator.total_labels:
            self.tree_page_label.config(
                text=f"Showing {start + 1}-{end} of {self.generator.total_labels}"
            )
        else:
            self.tree_page_label.config(text="No labels")
        self.tree_prev_button.config(state=tk.NORMAL if self._tree_page > 0 else tk.DISABLED)
        self.tree_next_button.config(
            state=tk.NORMAL if self._tree_page < page_count - 1 else tk.DISABLED
        )

        # Update count
        self.labels_count_label.config(
            text=f"Labels: {self.generator.total_labels} | Pages: {self.generator.total_pages}"
        )

        # Update spinbox range
        total_pages = max(1, self.generator.total_pages)
        self.page_spinbox.config(to=total_pages)
        self.total_pages_label.config(text=f"of {total_pages}")

    def _remove_selected_label(self):
        """Remove the selected label from the list."""
        selection = self.labels_tree.selection()
        if not selection:
            return

        indices = sorted([int(item) for item in selection if item.isdigit()], reverse=True)
        for idx in indices:
            if idx < len(self.generator.labels):
                del self.generator.labels[idx]

        # Removals shift positions, so any pending edit no longer refers to the
        # label it was opened on.
        self._editing_index = None

        self._update_labels_tree()
        self._update_status(f"Removed {len(indices)} label(s)")
        self._refresh_preview_if_visible()

    def _duplicate_selected_label(self):
        """Duplicate the selected labels."""
        selection = self.labels_tree.selection()
        if not selection:
            return

        indices = sorted([int(item) for item in selection if item.isdigit()])
        new_labels = []
        for idx in indices:
            if idx < len(self.generator.labels):
                label = self.generator.labels[idx]
                new_labels.append(
                    Label(
                        location_line1=label.location_line1,
                        location_line2=label.location_line2,
                        code=label.code,
                        date=label.date,
                        additional_info=label.additional_info,
                    )
                )

        self.generator.add_labels(new_labels)
        self._update_labels_tree()
        self._update_status(f"Duplicated {len(new_labels)} label(s)")
        self._refresh_preview_if_visible()

    def _edit_selected_label(self):
        """Edit the selected label."""
        selection = self.labels_tree.selection()
        if not selection or not selection[0].isdigit():
            return

        idx = int(selection[0])
        if idx >= len(self.generator.labels):
            return

        label = self.generator.labels[idx]

        # Fill form with label data
        self.entry_vars["location1"].set(label.location_line1)
        self.entry_vars["location2"].set(label.location_line2)
        self.entry_vars["code"].set(label.code)
        self.entry_vars["date"].set(label.date)
        self.entry_vars["notes"].set(label.additional_info)
        self.entry_vars["quantity"].set("1")

        # The label stays in the list until the edit is saved, so abandoning the
        # form (or closing the app) cannot lose it.
        self._editing_index = idx
        self._update_status("Editing label - press 'Add Label' to save changes")

    def _clear_labels(self):
        """Clear all labels."""
        if self.generator.labels:
            if messagebox.askyesno("Confirm", "Are you sure you want to remove all labels?"):
                self.generator.clear_labels()
                self._editing_index = None
                self._update_labels_tree()
                self._update_status("All labels cleared")
                self._update_preview()

    def _apply_config(self):
        """Apply configuration changes with validation."""
        try:
            # Validation with min and max bounds
            def get_int(name, min_val=1, max_val=None):
                val = int(self.config_vars[name].get())
                if val < min_val:
                    raise ValueError(f"{name} must be at least {min_val}")
                if max_val is not None and val > max_val:
                    raise ValueError(f"{name} must be at most {max_val}")
                return val

            def get_float(name, min_val=0.0, max_val=None):
                val = float(self.config_vars[name].get())
                if val < min_val:
                    raise ValueError(f"{name} must be at least {min_val}")
                if max_val is not None and val > max_val:
                    raise ValueError(f"{name} must be at most {max_val}")
                return val

            config = LabelConfig(
                labels_per_row=get_int("labels_per_row", 1, 50),
                labels_per_column=get_int("labels_per_column", 1, 50),
                label_width_mm=get_float("label_width_mm", LABEL_WIDTH_MM_MIN, LABEL_WIDTH_MM_MAX),
                label_height_mm=get_float(
                    "label_height_mm", LABEL_HEIGHT_MM_MIN, LABEL_HEIGHT_MM_MAX
                ),
                page_width_mm=get_float("page_width_mm", 10.0, 500.0),
                page_height_mm=get_float("page_height_mm", 10.0, 500.0),
                margin_top_mm=get_float("margin_top_mm", MARGIN_MM_MIN, MARGIN_MM_MAX),
                margin_bottom_mm=get_float("margin_bottom_mm", MARGIN_MM_MIN, MARGIN_MM_MAX),
                margin_left_mm=get_float("margin_left_mm", MARGIN_MM_MIN, MARGIN_MM_MAX),
                margin_right_mm=get_float("margin_right_mm", MARGIN_MM_MIN, MARGIN_MM_MAX),
                font_family=self.config_vars["font_family"].get(),
                font_size_pt=get_float("font_size_pt", FONT_SIZE_PT_MIN, FONT_SIZE_PT_MAX),
                line_spacing=get_float("line_spacing", 0.1, 5.0),
            )
            self.generator.config = config
            self._update_labels_tree()
            self._update_status("Configuration applied")
            self._update_preview()
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid value in configuration:\n{str(e)}")

    def _apply_preset(self, preset_name: str):
        """Apply a preset configuration."""
        presets = {
            "a4_standard": {
                "labels_per_row": "10",
                "labels_per_column": "13",
                "label_width_mm": "29.0",
                "label_height_mm": "13.0",
                "page_width_mm": "297",
                "page_height_mm": "210",
            },
            "a4_compact": {
                "labels_per_row": "12",
                "labels_per_column": "16",
                "label_width_mm": "24.0",
                "label_height_mm": "13.0",
                "page_width_mm": "297",
                "page_height_mm": "210",
            },
            "letter_us": {
                "labels_per_row": "10",
                "labels_per_column": "13",
                "label_width_mm": "27.0",
                "label_height_mm": "13.0",
                "page_width_mm": "279.4",
                "page_height_mm": "215.9",
            },
        }

        if preset_name in presets:
            for key, value in presets[preset_name].items():
                if key in self.config_vars:
                    self.config_vars[key].set(value)
            self._apply_config()

    def _update_preview(self):
        """Update the visual mockup preview."""
        # Clear current preview
        for widget in self.paper_frame.winfo_children():
            widget.destroy()

        if not self.generator.labels:
            lbl = ttk.Label(self.paper_frame, text="No labels to preview.", padding=50)
            lbl.pack()
            return

        try:
            page_num = int(self.page_var.get()) - 1
        except ValueError:
            page_num = 0

        if page_num < 0:
            page_num = 0
        if page_num >= self.generator.total_pages:
            page_num = max(0, self.generator.total_pages - 1)
            self.page_var.set(str(page_num + 1))

        grid = self.generator.get_labels_grid(page_num)

        # Display as a grid in the paper_frame
        config = self.generator.config

        # Scale the preview so a page fits on screen
        scale = PREVIEW_SCALE_FACTOR

        self.paper_frame.config(
            width=config.page_width_mm * scale, height=config.page_height_mm * scale, bg="white"
        )

        for r, row_labels in enumerate(grid):
            for c, label in enumerate(row_labels):
                if label:
                    # Create a "label" box
                    l_frame = tk.Frame(
                        self.paper_frame,
                        width=config.label_width_mm * scale,
                        height=config.label_height_mm * scale,
                        bg="white",
                        highlightbackground="#eee",
                        highlightthickness=1,
                    )
                    l_frame.grid(row=r, column=c)
                    # pack_propagate, not grid_propagate: the content frame
                    # below is packed, and grid_propagate only governs
                    # grid-managed children. With the wrong one the frame
                    # shrink-wraps its text and the mm dimensions above are
                    # silently discarded.
                    l_frame.pack_propagate(False)

                    # Add content
                    font_px = _preview_font_px(config, scale)
                    pad = max(1, int(round(config.label_padding_mm * scale)))
                    content_frame = tk.Frame(l_frame, bg="white")
                    content_frame.pack(fill=tk.BOTH, expand=True, padx=pad, pady=pad)

                    for line in render_label_lines(label, config):
                        if line.style == STYLE_SPACER:
                            tk.Label(
                                content_frame,
                                text="",
                                font=(config.font_family, -max(1, font_px // 2)),
                                bg="white",
                                **_PREVIEW_LINE_KW,
                            ).pack()  # Spacer
                            continue
                        style = ("italic",) if line.italic else ()
                        tk.Label(
                            content_frame,
                            text=line.text,
                            font=(
                                config.font_family,
                                -max(1, int(round(font_px * line.scale))),
                                *style,
                            ),
                            bg="white",
                            anchor="w",
                            **_PREVIEW_LINE_KW,
                        ).pack(fill=tk.X)
                else:
                    # Empty cell
                    l_frame = tk.Frame(
                        self.paper_frame,
                        width=config.label_width_mm * scale,
                        height=config.label_height_mm * scale,
                        bg="#fafafa",
                        highlightbackground="#f0f0f0",
                        highlightthickness=1,
                    )
                    l_frame.grid(row=r, column=c)

        # Update scrollregion
        self.root.update_idletasks()
        self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))

    def _open_in_browser(self):
        """Open the preview in the default browser."""
        if not self.generator.labels:
            messagebox.showinfo("Info", "No labels to display.")
            return

        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".html", delete=False, encoding="utf-8"
        ) as f:
            html = generate_html(self.generator)
            f.write(html)
            # Safely open the temp file in browser
            webbrowser.open(Path(f.name).resolve().as_uri())

    def _export(self, format_type: str):
        """Export labels to the specified format."""
        if not self.generator.labels:
            messagebox.showinfo("Info", "No labels to export.")
            return

        filetypes = {
            "html": [("HTML File", "*.html")],
            "pdf": [("PDF Document", "*.pdf")],
            "docx": [("Word Document", "*.docx")],
        }

        default_ext = {
            "html": ".html",
            "pdf": ".pdf",
            "docx": ".docx",
        }

        file_path = filedialog.asksaveasfilename(
            title=f"Export as {format_type.upper()}",
            filetypes=filetypes[format_type],
            defaultextension=default_ext[format_type],
        )

        if not file_path:
            return

        generators = {
            "html": generate_html,
            "pdf": generate_pdf,
            "docx": generate_docx,
        }

        def on_success(_result):
            self._update_status(f"Exported to {Path(file_path).name}")
            logger.info(f"Successfully exported to {file_path}")

            if messagebox.askyesno(
                "Export Successful", f"File saved to:\n{file_path}\n\nWould you like to open it?"
            ):
                # Safely open the file in browser/default app
                webbrowser.open(Path(file_path).resolve().as_uri())

        def on_error(exc):
            if isinstance(exc, ImportError):
                logger.error(f"Missing dependency for export: {exc}")
                messagebox.showerror("Missing Dependency", str(exc))
            else:
                logger.error(f"Export failed: {exc}", exc_info=exc)
                messagebox.showerror("Export Error", f"Failed to export:\n{exc}")
            self._update_status("Export failed")

        self._run_in_background(
            lambda: generators[format_type](self.generator, file_path),
            on_success,
            on_error,
            status=f"Exporting {format_type.upper()}...",
        )

    def _show_sequential_dialog(self):
        """Show dialog for generating sequential labels."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Generate Sequential Labels")
        dialog.geometry("450x400")
        dialog.transient(self.root)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)

        fields = [
            ("Location (Line 1):", "location1", ""),
            ("Location (Line 2):", "location2", ""),
            ("Code Prefix:", "prefix", "N"),
            ("Start Number:", "start", "1"),
            ("End Number:", "end", "10"),
            ("Date:", "date", ""),
        ]

        vars = {}
        for i, (label_text, var_name, default) in enumerate(fields):
            ttk.Label(frame, text=label_text).grid(row=i, column=0, sticky=tk.W, pady=8)
            var = tk.StringVar(value=default)
            vars[var_name] = var
            ttk.Entry(frame, textvariable=var, width=30).grid(
                row=i, column=1, sticky=tk.EW, pady=8, padx=(10, 0)
            )

        frame.columnconfigure(1, weight=1)

        def generate():
            try:
                labels = self.generator.generate_sequential_labels(
                    location_line1=vars["location1"].get(),
                    location_line2=vars["location2"].get(),
                    code_prefix=vars["prefix"].get(),
                    start_number=int(vars["start"].get()),
                    end_number=int(vars["end"].get()),
                    date=vars["date"].get(),
                )
                self.generator.add_labels(labels)
                self._update_labels_tree()
                self._update_status(f"Generated {len(labels)} sequential labels")
                dialog.destroy()
            except ValueError as e:
                messagebox.showerror("Input Error", f"Invalid numeric values:\n{str(e)}")

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=25)

        ttk.Button(btn_frame, text="Generate", command=generate).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def _show_help(self):
        """Show help dialog."""
        help_text = """Entomology Labels Generator - Guide

1. ADDING LABELS:
- Use the 'Add Single Label' form for individual entries.
- Use 'Import from File' to load data from Excel, CSV, Word, etc.
- Use 'Generate Sequential' for series (e.g., N1 to N100).

2. CONFIGURATION:
- Adjust label dimensions and page layout in the 'Configuration' tab.
- Presets are available for standard A4 and US Letter sizes.

3. EXPORTING:
- HTML: Great for quick printing from a browser.
- PDF: Best for preserving exact dimensions (requires weasyprint).
- DOCX: Use if you need to manually edit labels in Word.

TIPS:
- When printing HTML/PDF, set margins to 'None' in the print dialog.
- The visual preview shows one page at a time.
"""
        messagebox.showinfo("Guide", help_text)

    def _show_about(self):
        """Show about dialog."""
        about_text = """Entomology Labels Generator
Version 1.2.0

A professional tool for biological specimen labeling.

Developed for entomologists and museum curators.
License: MIT
"""
        messagebox.showinfo("About", about_text)

    def _update_status(self, message: str):
        """Update status bar message."""
        self.status_label.config(text=message)

    def run(self):
        """Run the GUI application."""
        self.root.mainloop()


def main():
    """Entry point for the GUI application."""
    app = EntomologyLabelsGUI()
    app.run()


if __name__ == "__main__":
    main()
