"""
Graphical User Interface for Entomology Labels Generator.

Provides an easy-to-use interface for creating and exporting entomology labels.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
from typing import Optional
import webbrowser
import tempfile
import json

from .label_generator import LabelGenerator, Label, LabelConfig
from .input_handlers import load_data
from .output_generators import generate_html, generate_pdf, generate_docx


class EntomologyLabelsGUI:
    """Main GUI application for generating entomology labels."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Entomology Labels Generator")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)

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
        file_menu.add_command(label="Importa dati...", command=self._import_data, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="Esporta HTML...", command=lambda: self._export("html"))
        file_menu.add_command(label="Esporta PDF...", command=lambda: self._export("pdf"))
        file_menu.add_command(label="Esporta DOCX...", command=lambda: self._export("docx"))
        file_menu.add_separator()
        file_menu.add_command(label="Esci", command=self.root.quit, accelerator="Ctrl+Q")

        # Edit menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Modifica", menu=edit_menu)
        edit_menu.add_command(label="Cancella tutte le etichette", command=self._clear_labels)
        edit_menu.add_command(label="Genera etichette sequenziali...", command=self._show_sequential_dialog)

        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Aiuto", menu=help_menu)
        help_menu.add_command(label="Guida", command=self._show_help)
        help_menu.add_command(label="Informazioni", command=self._show_about)

    def _setup_main_layout(self):
        """Setup the main application layout."""
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Tab 1: Data Entry
        self._setup_data_tab()

        # Tab 2: Configuration
        self._setup_config_tab()

        # Tab 3: Preview
        self._setup_preview_tab()

        # Bottom status bar
        self._setup_status_bar(main_frame)

    def _setup_data_tab(self):
        """Setup the data entry tab."""
        data_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(data_frame, text="Dati Etichette")

        # Left panel - Form for single label entry
        left_frame = ttk.LabelFrame(data_frame, text="Aggiungi Etichetta", padding="10")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        # Form fields
        fields = [
            ("Località (riga 1):", "location1"),
            ("Località (riga 2):", "location2"),
            ("Codice:", "code"),
            ("Data:", "date"),
            ("Note aggiuntive:", "notes"),
            ("Quantità:", "quantity"),
        ]

        self.entry_vars = {}
        for i, (label, var_name) in enumerate(fields):
            ttk.Label(left_frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=2)
            var = tk.StringVar()
            self.entry_vars[var_name] = var
            if var_name == "quantity":
                var.set("1")
            entry = ttk.Entry(left_frame, textvariable=var, width=40)
            entry.grid(row=i, column=1, sticky=tk.EW, pady=2, padx=(5, 0))

        left_frame.columnconfigure(1, weight=1)

        # Buttons
        btn_frame = ttk.Frame(left_frame)
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=10)

        ttk.Button(btn_frame, text="Aggiungi", command=self._add_label).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Pulisci campi", command=self._clear_form).pack(side=tk.LEFT, padx=2)

        # Import button
        import_frame = ttk.LabelFrame(left_frame, text="Importa da file", padding="10")
        import_frame.grid(row=len(fields)+1, column=0, columnspan=2, sticky=tk.EW, pady=10)

        ttk.Button(import_frame, text="Importa (Excel, CSV, TXT, DOCX, JSON)...",
                   command=self._import_data).pack(fill=tk.X)

        ttk.Label(import_frame, text="Formati supportati: .xlsx, .xls, .csv, .txt, .docx, .json, .yaml",
                  foreground="gray").pack(pady=(5, 0))

        # Right panel - Labels list
        right_frame = ttk.LabelFrame(data_frame, text="Etichette Caricate", padding="10")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

        # Treeview for labels
        columns = ("location1", "location2", "code", "date")
        self.labels_tree = ttk.Treeview(right_frame, columns=columns, show="headings", height=15)

        self.labels_tree.heading("location1", text="Località 1")
        self.labels_tree.heading("location2", text="Località 2")
        self.labels_tree.heading("code", text="Codice")
        self.labels_tree.heading("date", text="Data")

        self.labels_tree.column("location1", width=150)
        self.labels_tree.column("location2", width=150)
        self.labels_tree.column("code", width=80)
        self.labels_tree.column("date", width=100)

        # Scrollbar
        scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.labels_tree.yview)
        self.labels_tree.configure(yscrollcommand=scrollbar.set)

        self.labels_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Buttons under treeview
        tree_btn_frame = ttk.Frame(right_frame)
        tree_btn_frame.pack(fill=tk.X, pady=(5, 0))

        ttk.Button(tree_btn_frame, text="Rimuovi selezionata",
                   command=self._remove_selected_label).pack(side=tk.LEFT, padx=2)
        ttk.Button(tree_btn_frame, text="Rimuovi tutte",
                   command=self._clear_labels).pack(side=tk.LEFT, padx=2)

    def _setup_config_tab(self):
        """Setup the configuration tab."""
        config_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(config_frame, text="Configurazione")

        # Layout settings
        layout_frame = ttk.LabelFrame(config_frame, text="Layout Pagina", padding="10")
        layout_frame.pack(fill=tk.X, pady=5)

        layout_fields = [
            ("Etichette per riga:", "labels_per_row", "10"),
            ("Etichette per colonna:", "labels_per_column", "13"),
            ("Larghezza etichetta (mm):", "label_width_mm", "29.0"),
            ("Altezza etichetta (mm):", "label_height_mm", "13.0"),
        ]

        self.config_vars = {}
        for i, (label, var_name, default) in enumerate(layout_fields):
            ttk.Label(layout_frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=2)
            var = tk.StringVar(value=default)
            self.config_vars[var_name] = var
            entry = ttk.Entry(layout_frame, textvariable=var, width=15)
            entry.grid(row=i, column=1, sticky=tk.W, pady=2, padx=(5, 0))

        # Page settings
        page_frame = ttk.LabelFrame(config_frame, text="Dimensioni Pagina", padding="10")
        page_frame.pack(fill=tk.X, pady=5)

        page_fields = [
            ("Larghezza pagina (mm):", "page_width_mm", "297"),
            ("Altezza pagina (mm):", "page_height_mm", "210"),
            ("Margine superiore (mm):", "margin_top_mm", "0"),
            ("Margine inferiore (mm):", "margin_bottom_mm", "0"),
            ("Margine sinistro (mm):", "margin_left_mm", "0"),
            ("Margine destro (mm):", "margin_right_mm", "0"),
        ]

        for i, (label, var_name, default) in enumerate(page_fields):
            ttk.Label(page_frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=2)
            var = tk.StringVar(value=default)
            self.config_vars[var_name] = var
            entry = ttk.Entry(page_frame, textvariable=var, width=15)
            entry.grid(row=i, column=1, sticky=tk.W, pady=2, padx=(5, 0))

        # Font settings
        font_frame = ttk.LabelFrame(config_frame, text="Carattere", padding="10")
        font_frame.pack(fill=tk.X, pady=5)

        ttk.Label(font_frame, text="Font:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.config_vars["font_family"] = tk.StringVar(value="Arial")
        font_combo = ttk.Combobox(font_frame, textvariable=self.config_vars["font_family"],
                                   values=["Arial", "Times New Roman", "Helvetica", "Calibri", "Courier New"],
                                   width=20)
        font_combo.grid(row=0, column=1, sticky=tk.W, pady=2, padx=(5, 0))

        ttk.Label(font_frame, text="Dimensione (pt):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.config_vars["font_size_pt"] = tk.StringVar(value="6")
        ttk.Entry(font_frame, textvariable=self.config_vars["font_size_pt"], width=15).grid(
            row=1, column=1, sticky=tk.W, pady=2, padx=(5, 0))

        ttk.Label(font_frame, text="Interlinea:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.config_vars["line_spacing"] = tk.StringVar(value="1.0")
        ttk.Entry(font_frame, textvariable=self.config_vars["line_spacing"], width=15).grid(
            row=2, column=1, sticky=tk.W, pady=2, padx=(5, 0))

        # Apply button
        ttk.Button(config_frame, text="Applica configurazione",
                   command=self._apply_config).pack(pady=10)

        # Presets
        preset_frame = ttk.LabelFrame(config_frame, text="Preimpostazioni", padding="10")
        preset_frame.pack(fill=tk.X, pady=5)

        ttk.Button(preset_frame, text="A4 Orizzontale (10x13)",
                   command=lambda: self._apply_preset("a4_standard")).pack(side=tk.LEFT, padx=2)
        ttk.Button(preset_frame, text="A4 Compatto (12x15)",
                   command=lambda: self._apply_preset("a4_compact")).pack(side=tk.LEFT, padx=2)
        ttk.Button(preset_frame, text="Letter US Orizzontale (10x12)",
                   command=lambda: self._apply_preset("letter_us")).pack(side=tk.LEFT, padx=2)

    def _setup_preview_tab(self):
        """Setup the preview tab."""
        preview_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(preview_frame, text="Anteprima & Esporta")

        # Preview controls
        controls_frame = ttk.Frame(preview_frame)
        controls_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(controls_frame, text="Aggiorna anteprima",
                   command=self._update_preview).pack(side=tk.LEFT, padx=2)
        ttk.Button(controls_frame, text="Apri in browser",
                   command=self._open_in_browser).pack(side=tk.LEFT, padx=2)

        ttk.Separator(controls_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Button(controls_frame, text="Esporta HTML",
                   command=lambda: self._export("html")).pack(side=tk.LEFT, padx=2)
        ttk.Button(controls_frame, text="Esporta PDF",
                   command=lambda: self._export("pdf")).pack(side=tk.LEFT, padx=2)
        ttk.Button(controls_frame, text="Esporta DOCX",
                   command=lambda: self._export("docx")).pack(side=tk.LEFT, padx=2)

        # Preview info
        self.preview_info = ttk.Label(preview_frame, text="")
        self.preview_info.pack(fill=tk.X)

        # Preview text (HTML source)
        preview_text_frame = ttk.LabelFrame(preview_frame, text="Anteprima HTML", padding="5")
        preview_text_frame.pack(fill=tk.BOTH, expand=True)

        self.preview_text = scrolledtext.ScrolledText(preview_text_frame, wrap=tk.WORD, height=20)
        self.preview_text.pack(fill=tk.BOTH, expand=True)

    def _setup_status_bar(self, parent):
        """Setup the status bar."""
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill=tk.X, pady=(10, 0))

        self.status_label = ttk.Label(status_frame, text="Pronto")
        self.status_label.pack(side=tk.LEFT)

        self.labels_count_label = ttk.Label(status_frame, text="Etichette: 0 | Pagine: 0")
        self.labels_count_label.pack(side=tk.RIGHT)

    def _setup_bindings(self):
        """Setup keyboard bindings."""
        self.root.bind("<Control-o>", lambda e: self._import_data())
        self.root.bind("<Control-q>", lambda e: self.root.quit())

    def _add_label(self):
        """Add a label from the form."""
        try:
            quantity = int(self.entry_vars["quantity"].get() or "1")
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
            messagebox.showwarning("Attenzione", "Inserire almeno un campo per l'etichetta.")
            return

        for _ in range(quantity):
            self.generator.add_label(Label(
                location_line1=label.location_line1,
                location_line2=label.location_line2,
                code=label.code,
                date=label.date,
                additional_info=label.additional_info,
            ))

        self._update_labels_tree()
        self._clear_form()
        self._update_status(f"Aggiunte {quantity} etichette")

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
            ("Tutti i formati supportati", "*.xlsx *.xls *.csv *.txt *.docx *.json *.yaml *.yml"),
            ("Excel", "*.xlsx *.xls"),
            ("CSV", "*.csv"),
            ("Testo", "*.txt"),
            ("Word", "*.docx"),
            ("JSON", "*.json"),
            ("YAML", "*.yaml *.yml"),
        ]

        file_path = filedialog.askopenfilename(
            title="Seleziona file da importare",
            filetypes=filetypes
        )

        if not file_path:
            return

        try:
            labels = load_data(file_path)
            self.generator.add_labels(labels)
            self._update_labels_tree()
            self._update_status(f"Importate {len(labels)} etichette da {Path(file_path).name}")
        except Exception as e:
            messagebox.showerror("Errore importazione", f"Errore durante l'importazione:\n{str(e)}")

    def _update_labels_tree(self):
        """Update the labels treeview."""
        # Clear existing items
        for item in self.labels_tree.get_children():
            self.labels_tree.delete(item)

        # Add all labels
        for i, label in enumerate(self.generator.labels):
            self.labels_tree.insert("", tk.END, iid=str(i), values=(
                label.location_line1[:30] + "..." if len(label.location_line1) > 30 else label.location_line1,
                label.location_line2[:30] + "..." if len(label.location_line2) > 30 else label.location_line2,
                label.code,
                label.date,
            ))

        # Update count
        self.labels_count_label.config(
            text=f"Etichette: {self.generator.total_labels} | Pagine: {self.generator.total_pages}"
        )

    def _remove_selected_label(self):
        """Remove the selected label from the list."""
        selection = self.labels_tree.selection()
        if not selection:
            return

        indices = sorted([int(item) for item in selection], reverse=True)
        for idx in indices:
            if idx < len(self.generator.labels):
                del self.generator.labels[idx]

        self._update_labels_tree()
        self._update_status(f"Rimosse {len(indices)} etichette")

    def _clear_labels(self):
        """Clear all labels."""
        if self.generator.labels:
            if messagebox.askyesno("Conferma", "Vuoi rimuovere tutte le etichette?"):
                self.generator.clear_labels()
                self._update_labels_tree()
                self._update_status("Tutte le etichette rimosse")

    def _apply_config(self):
        """Apply configuration changes."""
        try:
            config = LabelConfig(
                labels_per_row=int(self.config_vars["labels_per_row"].get()),
                labels_per_column=int(self.config_vars["labels_per_column"].get()),
                label_width_mm=float(self.config_vars["label_width_mm"].get()),
                label_height_mm=float(self.config_vars["label_height_mm"].get()),
                page_width_mm=float(self.config_vars["page_width_mm"].get()),
                page_height_mm=float(self.config_vars["page_height_mm"].get()),
                margin_top_mm=float(self.config_vars["margin_top_mm"].get()),
                margin_bottom_mm=float(self.config_vars["margin_bottom_mm"].get()),
                margin_left_mm=float(self.config_vars["margin_left_mm"].get()),
                margin_right_mm=float(self.config_vars["margin_right_mm"].get()),
                font_family=self.config_vars["font_family"].get(),
                font_size_pt=float(self.config_vars["font_size_pt"].get()),
                line_spacing=float(self.config_vars["line_spacing"].get()),
            )
            self.generator.config = config
            self._update_labels_tree()
            self._update_status("Configurazione applicata")
        except ValueError as e:
            messagebox.showerror("Errore", f"Valore non valido nella configurazione:\n{str(e)}")

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
        """Update the HTML preview."""
        if not self.generator.labels:
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert("1.0", "Nessuna etichetta da visualizzare.\n\nAggiungi etichette dalla scheda 'Dati Etichette'.")
            self.preview_info.config(text="")
            return

        html = generate_html(self.generator)
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert("1.0", html)

        self.preview_info.config(
            text=f"Totale: {self.generator.total_labels} etichette su {self.generator.total_pages} pagine"
        )

    def _open_in_browser(self):
        """Open the preview in the default browser."""
        if not self.generator.labels:
            messagebox.showinfo("Info", "Nessuna etichetta da visualizzare.")
            return

        with tempfile.NamedTemporaryFile(mode="w", suffix=".html", delete=False, encoding="utf-8") as f:
            html = generate_html(self.generator)
            f.write(html)
            webbrowser.open(f"file://{f.name}")

    def _export(self, format_type: str):
        """Export labels to the specified format."""
        if not self.generator.labels:
            messagebox.showinfo("Info", "Nessuna etichetta da esportare.")
            return

        filetypes = {
            "html": [("HTML", "*.html")],
            "pdf": [("PDF", "*.pdf")],
            "docx": [("Word Document", "*.docx")],
        }

        default_ext = {
            "html": ".html",
            "pdf": ".pdf",
            "docx": ".docx",
        }

        file_path = filedialog.asksaveasfilename(
            title=f"Salva come {format_type.upper()}",
            filetypes=filetypes[format_type],
            defaultextension=default_ext[format_type]
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

            self._update_status(f"Esportato in {Path(file_path).name}")

            if messagebox.askyesno("Esportazione completata",
                                    f"File salvato in:\n{file_path}\n\nVuoi aprirlo?"):
                webbrowser.open(f"file://{file_path}")

        except ImportError as e:
            messagebox.showerror("Dipendenza mancante", str(e))
        except Exception as e:
            messagebox.showerror("Errore esportazione", f"Errore durante l'esportazione:\n{str(e)}")

    def _show_sequential_dialog(self):
        """Show dialog for generating sequential labels."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Genera etichette sequenziali")
        dialog.geometry("400x350")
        dialog.transient(self.root)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)

        fields = [
            ("Località (riga 1):", "location1", "Italia, Trentino Alto Adige,"),
            ("Località (riga 2):", "location2", "Giustino (TN), Vedretta d'Amola"),
            ("Prefisso codice:", "prefix", "N"),
            ("Numero iniziale:", "start", "1"),
            ("Numero finale:", "end", "10"),
            ("Data:", "date", ""),
        ]

        vars = {}
        for i, (label, var_name, default) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=5)
            var = tk.StringVar(value=default)
            vars[var_name] = var
            ttk.Entry(frame, textvariable=var, width=30).grid(row=i, column=1, sticky=tk.EW, pady=5, padx=(5, 0))

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
                self._update_status(f"Generate {len(labels)} etichette sequenziali")
                dialog.destroy()
            except ValueError as e:
                messagebox.showerror("Errore", f"Valori non validi:\n{str(e)}")

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=20)

        ttk.Button(btn_frame, text="Genera", command=generate).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Annulla", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def _show_help(self):
        """Show help dialog."""
        help_text = """Entomology Labels Generator - Guida

IMPORTAZIONE DATI:
- Supporta Excel (.xlsx, .xls), CSV, TXT, Word (.docx), JSON, YAML
- Le colonne devono contenere: location_line1, location_line2, code, date
- Usa la colonna "count" per duplicare le etichette

CONFIGURAZIONE:
- 10 etichette per riga x 13 per colonna (default A4)
- Personalizza dimensioni, margini e font

ESPORTAZIONE:
- HTML: Apri in browser e usa "Stampa > Salva come PDF"
- PDF: Richiede weasyprint installato
- DOCX: Per modifiche in Microsoft Word

FORMATO ETICHETTA:
Riga 1: Località principale
Riga 2: Località secondaria
[riga vuota]
Codice specimen
Data raccolta
"""
        messagebox.showinfo("Guida", help_text)

    def _show_about(self):
        """Show about dialog."""
        about_text = """Entomology Labels Generator
Versione 1.0.0

Genera etichette professionali per specimen entomologici.

Supporta:
- Importazione: Excel, CSV, TXT, Word, JSON, YAML
- Esportazione: HTML, PDF, DOCX

Licenza: MIT
"""
        messagebox.showinfo("Informazioni", about_text)

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
