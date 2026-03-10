"""
Graphical User Interface for Entomology Labels Generator.

Provides an easy-to-use interface for creating and exporting entomology labels.
"""

import json
import tempfile
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Optional

from .input_handlers import load_data
from .label_generator import Label, LabelConfig, LabelGenerator
from .output_generators import generate_docx, generate_html, generate_pdf


class EntomologyLabelsGUI:
    """Main GUI application for generating entomology labels."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Entomology Labels Generator")
        self.root.geometry("1100x800")
        self.root.minsize(900, 700)

        # Initialize generator
        self.generator = LabelGenerator()

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
        self.labels_tree.heading("quantity", text="Qty")

        self.labels_tree.column("location1", width=150)
        self.labels_tree.column("location2", width=150)
        self.labels_tree.column("code", width=70)
        self.labels_tree.column("date", width=90)
        self.labels_tree.column("quantity", width=40)

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

        # Handle mouse wheel scrolling
        def _on_mousewheel(event):
            config_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        config_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        # Note: we need to pack scrollbar somewhere, but notebook makes it tricky.
        # Let's use a simpler layout for now if scroll isn't strictly needed,
        # but let's stick to a structured layout.

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

        for _ in range(quantity):
            self.generator.add_label(
                Label(
                    location_line1=label.location_line1,
                    location_line2=label.location_line2,
                    code=label.code,
                    date=label.date,
                    additional_info=label.additional_info,
                )
            )

        self._update_labels_tree()
        self._clear_form()
        self._update_status(f"Added {quantity} label(s)")

        # Trigger preview update if on preview tab
        if self.notebook.index(self.notebook.select()) == 1:
            self._update_preview()

    def _clear_form(self):
        """Clear the entry form."""
        for var_name, var in self.entry_vars.items():
            if var_name == "quantity":
                var.set("1")
            else:
                var.set("")

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

        try:
            labels = load_data(file_path)
            self.generator.add_labels(labels)
            self._update_labels_tree()
            self._update_status(f"Imported {len(labels)} labels from {Path(file_path).name}")

            # Switch to data tab and update
            self.notebook.select(0)

        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import data:\n{str(e)}")

    def _update_labels_tree(self):
        """Update the labels treeview."""
        # Clear existing items
        for item in self.labels_tree.get_children():
            self.labels_tree.delete(item)

        # Add labels, grouped for display
        # We'll just show the actual labels for now as they are in the generator
        # To make it more readable, we could group identical labels, but let's keep it simple
        for i, label in enumerate(self.generator.labels):
            # Only show first 500 to prevent GUI lag
            if i >= 500:
                self.labels_tree.insert("", tk.END, values=("...", "... and more ...", "", "", ""))
                break

            self.labels_tree.insert(
                "",
                tk.END,
                iid=str(i),
                values=(
                    label.location_line1[:30] + ("..." if len(label.location_line1) > 30 else ""),
                    label.location_line2[:30] + ("..." if len(label.location_line2) > 30 else ""),
                    label.code,
                    label.date,
                    "1",
                ),
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

        self._update_labels_tree()
        self._update_status(f"Removed {len(indices)} label(s)")

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

    def _edit_selected_label(self):
        """Edit the selected label."""
        selection = self.labels_tree.selection()
        if not selection:
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

        # Remove it from the list (user will "add" it back after editing)
        del self.generator.labels[idx]
        self._update_labels_tree()
        self._update_status("Editing label (restored to form)")

    def _clear_labels(self):
        """Clear all labels."""
        if self.generator.labels:
            if messagebox.askyesno("Confirm", "Are you sure you want to remove all labels?"):
                self.generator.clear_labels()
                self._update_labels_tree()
                self._update_status("All labels cleared")
                self._update_preview()

    def _apply_config(self):
        """Apply configuration changes with validation."""
        try:
            # Basic validation
            def get_int(name, min_val=1):
                val = int(self.config_vars[name].get())
                if val < min_val:
                    raise ValueError(f"{name} must be at least {min_val}")
                return val

            def get_float(name, min_val=0.0):
                val = float(self.config_vars[name].get())
                if val < min_val:
                    raise ValueError(f"{name} must be at least {min_val}")
                return val

            config = LabelConfig(
                labels_per_row=get_int("labels_per_row"),
                labels_per_column=get_int("labels_per_column"),
                label_width_mm=get_float("label_width_mm", 1.0),
                label_height_mm=get_float("label_height_mm", 1.0),
                page_width_mm=get_float("page_width_mm", 10.0),
                page_height_mm=get_float("page_height_mm", 10.0),
                margin_top_mm=get_float("margin_top_mm"),
                margin_bottom_mm=get_float("margin_bottom_mm"),
                margin_left_mm=get_float("margin_left_mm"),
                margin_right_mm=get_float("margin_right_mm"),
                font_family=self.config_vars["font_family"].get(),
                font_size_pt=get_float("font_size_pt", 1.0),
                line_spacing=get_float("line_spacing", 0.1),
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

        # We'll use a scale for the preview so it fits on screen
        # 1mm = ~3 pixels for preview
        scale = 3.5

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
                    l_frame.grid_propagate(False)

                    # Add content
                    font_size = max(4, int(config.font_size_pt * scale / 3))
                    content_frame = tk.Frame(l_frame, bg="white")
                    content_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

                    tk.Label(
                        content_frame,
                        text=label.location_line1,
                        font=(config.font_family, font_size),
                        bg="white",
                        anchor="w",
                    ).pack(fill=tk.X)
                    tk.Label(
                        content_frame,
                        text=label.location_line2,
                        font=(config.font_family, font_size),
                        bg="white",
                        anchor="w",
                    ).pack(fill=tk.X)
                    tk.Label(
                        content_frame,
                        text="",
                        font=(config.font_family, font_size // 2),
                        bg="white",
                    ).pack()  # Spacer
                    tk.Label(
                        content_frame,
                        text=label.code,
                        font=(config.font_family, font_size),
                        bg="white",
                        anchor="w",
                    ).pack(fill=tk.X)
                    tk.Label(
                        content_frame,
                        text=label.date,
                        font=(config.font_family, font_size),
                        bg="white",
                        anchor="w",
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
            webbrowser.open(Path(f.name).absolute().as_uri())

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

        try:
            if format_type == "html":
                generate_html(self.generator, file_path)
            elif format_type == "pdf":
                generate_pdf(self.generator, file_path)
            elif format_type == "docx":
                generate_docx(self.generator, file_path)

            self._update_status(f"Exported to {Path(file_path).name}")

            if messagebox.askyesno(
                "Export Successful", f"File saved to:\n{file_path}\n\nWould you like to open it?"
            ):
                webbrowser.open(Path(file_path).absolute().as_uri())

        except ImportError as e:
            messagebox.showerror("Missing Dependency", str(e))
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export:\n{str(e)}")

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
