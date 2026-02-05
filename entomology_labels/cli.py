"""
Command Line Interface for Entomology Labels Generator.

Provides commands for generating labels from various input formats.
"""

import click
from pathlib import Path
from typing import Optional

from .label_generator import LabelGenerator, LabelConfig
from .input_handlers import load_data
from .output_generators import generate_html, generate_pdf, generate_docx


@click.group()
@click.version_option(version="1.0.0", prog_name="entomology-labels")
def cli():
    """Entomology Labels Generator - Create professional specimen labels.

    Generate entomology labels from various input formats (Excel, CSV, TXT, DOCX, JSON, YAML)
    and export them to HTML, PDF, or DOCX.

    Examples:

      # Generate labels from Excel to HTML
      entomology-labels generate data.xlsx -o labels.html

      # Generate labels to PDF with custom layout
      entomology-labels generate data.csv -o labels.pdf --rows 10 --cols 13

      # Launch the GUI
      entomology-labels gui
    """
    pass


@cli.command()
@click.argument("input_file", type=click.Path(exists=True))
@click.option("-o", "--output", required=True, help="Output file path (.html, .pdf, or .docx)")
@click.option("--rows", default=10, type=int, help="Labels per row (default: 10)")
@click.option("--cols", default=13, type=int, help="Labels per column (default: 13)")
@click.option("--label-width", default=21.0, type=float, help="Label width in mm (default: 21.0)")
@click.option("--label-height", default=22.85, type=float, help="Label height in mm (default: 22.85)")
@click.option("--page-width", default=210.0, type=float, help="Page width in mm (default: 210 for A4)")
@click.option("--page-height", default=297.0, type=float, help="Page height in mm (default: 297 for A4)")
@click.option("--font-size", default=6.0, type=float, help="Font size in points (default: 6)")
@click.option("--font-family", default="Arial", help="Font family (default: Arial)")
@click.option("--open", "open_after", is_flag=True, help="Open file after generation")
@click.option("-v", "--verbose", is_flag=True, help="Enable verbose output")
def generate(
    input_file: str,
    output: str,
    rows: int,
    cols: int,
    label_width: float,
    label_height: float,
    page_width: float,
    page_height: float,
    font_size: float,
    font_family: str,
    open_after: bool,
    verbose: bool
):
    """Generate labels from an input file.

    INPUT_FILE: Path to the input file (Excel, CSV, TXT, DOCX, JSON, or YAML)

    Examples:

      entomology-labels generate specimens.xlsx -o labels.html

      entomology-labels generate data.csv -o output.pdf --rows 12 --cols 15

      entomology-labels generate labels.json -o output.docx --open
    """
    input_path = Path(input_file)
    output_path = Path(output)

    # Determine output format
    output_format = output_path.suffix.lower()
    if output_format not in [".html", ".pdf", ".docx"]:
        raise click.ClickException(
            f"Unsupported output format: {output_format}. "
            "Use .html, .pdf, or .docx"
        )

    if verbose:
        click.echo(f"Loading data from: {input_path}")

    # Load data
    try:
        labels = load_data(input_path)
    except Exception as e:
        raise click.ClickException(f"Error loading data: {e}")

    if not labels:
        raise click.ClickException("No labels found in input file")

    if verbose:
        click.echo(f"Loaded {len(labels)} labels")

    # Configure generator
    config = LabelConfig(
        labels_per_row=rows,
        labels_per_column=cols,
        label_width_mm=label_width,
        label_height_mm=label_height,
        page_width_mm=page_width,
        page_height_mm=page_height,
        font_size_pt=font_size,
        font_family=font_family,
    )

    generator = LabelGenerator(config)
    generator.add_labels(labels)

    if verbose:
        click.echo(f"Configuration: {rows}x{cols} labels per page")
        click.echo(f"Total pages: {generator.total_pages}")

    # Generate output
    try:
        if output_format == ".html":
            generate_html(generator, output_path, open_in_browser=open_after)
        elif output_format == ".pdf":
            generate_pdf(generator, output_path, open_after=open_after)
        elif output_format == ".docx":
            generate_docx(generator, output_path, open_after=open_after)

        click.echo(f"Generated {generator.total_labels} labels on {generator.total_pages} pages")
        click.echo(f"Output saved to: {output_path}")

    except ImportError as e:
        raise click.ClickException(str(e))
    except Exception as e:
        raise click.ClickException(f"Error generating output: {e}")


@cli.command()
@click.option("--location1", required=True, help="First location line")
@click.option("--location2", required=True, help="Second location line")
@click.option("--prefix", required=True, help="Code prefix (e.g., 'N' for N1, N2...)")
@click.option("--start", required=True, type=int, help="Start number")
@click.option("--end", required=True, type=int, help="End number")
@click.option("--date", default="", help="Collection date")
@click.option("-o", "--output", required=True, help="Output file path")
@click.option("--rows", default=10, type=int, help="Labels per row")
@click.option("--cols", default=13, type=int, help="Labels per column")
@click.option("--open", "open_after", is_flag=True, help="Open file after generation")
def sequence(
    location1: str,
    location2: str,
    prefix: str,
    start: int,
    end: int,
    date: str,
    output: str,
    rows: int,
    cols: int,
    open_after: bool
):
    """Generate sequential labels with incrementing codes.

    Example:

      entomology-labels sequence \\
        --location1 "Italia, Trentino Alto Adige," \\
        --location2 "Giustino (TN), Vedretta d'Amola" \\
        --prefix N --start 1 --end 50 \\
        --date "15.vi.2024" \\
        -o labels.html
    """
    output_path = Path(output)
    output_format = output_path.suffix.lower()

    if output_format not in [".html", ".pdf", ".docx"]:
        raise click.ClickException(
            f"Unsupported output format: {output_format}. "
            "Use .html, .pdf, or .docx"
        )

    config = LabelConfig(labels_per_row=rows, labels_per_column=cols)
    generator = LabelGenerator(config)

    labels = generator.generate_sequential_labels(
        location_line1=location1,
        location_line2=location2,
        code_prefix=prefix,
        start_number=start,
        end_number=end,
        date=date,
    )
    generator.add_labels(labels)

    try:
        if output_format == ".html":
            generate_html(generator, output_path, open_in_browser=open_after)
        elif output_format == ".pdf":
            generate_pdf(generator, output_path, open_after=open_after)
        elif output_format == ".docx":
            generate_docx(generator, output_path, open_after=open_after)

        click.echo(f"Generated {len(labels)} sequential labels ({prefix}{start} to {prefix}{end})")
        click.echo(f"Output saved to: {output_path}")

    except ImportError as e:
        raise click.ClickException(str(e))
    except Exception as e:
        raise click.ClickException(f"Error generating output: {e}")


@cli.command()
def gui():
    """Launch the graphical user interface."""
    try:
        from .gui import main as gui_main
        gui_main()
    except ImportError as e:
        raise click.ClickException(
            f"GUI dependencies not available: {e}\n"
            "Make sure tkinter is installed."
        )


@cli.command()
@click.argument("output_file", type=click.Path())
@click.option("--format", "file_format", type=click.Choice(["json", "yaml", "csv", "excel"]),
              default="json", help="Template format")
def template(output_file: str, file_format: str):
    """Generate a template file for label data.

    Creates an example file that you can fill with your own data.

    Example:

      entomology-labels template my_labels.json

      entomology-labels template my_labels.xlsx --format excel
    """
    output_path = Path(output_file)

    example_data = [
        {
            "location_line1": "Italia, Trentino Alto Adige,",
            "location_line2": "Giustino (TN), Vedretta d'Amola",
            "code": "N1",
            "date": "15.vi.2024",
            "additional_info": "",
            "count": 1
        },
        {
            "location_line1": "Italia, Trentino Alto Adige,",
            "location_line2": "Giustino (TN), Vedretta d'Amola",
            "code": "N2",
            "date": "15.vi.2024",
            "additional_info": "",
            "count": 1
        },
        {
            "location_line1": "Italia, Lombardia,",
            "location_line2": "Sondrio, Valmalenco",
            "code": "H1",
            "date": "20.vii.2024",
            "additional_info": "leg. Rossi",
            "count": 3
        },
    ]

    try:
        if file_format == "json":
            import json
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump({"labels": example_data}, f, indent=2, ensure_ascii=False)

        elif file_format == "yaml":
            try:
                import yaml
            except ImportError:
                raise click.ClickException("PyYAML is required. Install with: pip install pyyaml")
            with open(output_path, "w", encoding="utf-8") as f:
                yaml.dump({"labels": example_data}, f, default_flow_style=False, allow_unicode=True)

        elif file_format == "csv":
            import csv
            with open(output_path, "w", encoding="utf-8", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=example_data[0].keys())
                writer.writeheader()
                writer.writerows(example_data)

        elif file_format == "excel":
            try:
                import pandas as pd
            except ImportError:
                raise click.ClickException("pandas and openpyxl are required. Install with: pip install pandas openpyxl")
            df = pd.DataFrame(example_data)
            df.to_excel(output_path, index=False)

        click.echo(f"Template created: {output_path}")
        click.echo("\nEdit this file with your data, then use:")
        click.echo(f"  entomology-labels generate {output_path} -o labels.html")

    except Exception as e:
        raise click.ClickException(f"Error creating template: {e}")


@cli.command()
def info():
    """Show information about supported formats and configuration."""
    info_text = """
Entomology Labels Generator
===========================

SUPPORTED INPUT FORMATS:
  - Excel (.xlsx, .xls)
  - CSV (.csv)
  - Text (.txt) - tab-separated or key-value pairs
  - Word (.docx) - table or paragraph format
  - JSON (.json)
  - YAML (.yaml, .yml)

EXPECTED COLUMNS/FIELDS:
  - location_line1 (or location1, località1)
  - location_line2 (or location2, località2)
  - code (or specimen_code, codice)
  - date (or collection_date, data)
  - additional_info (or notes, note) - optional
  - count (or quantity) - optional, for duplicating labels

OUTPUT FORMATS:
  - HTML (.html) - Open in browser, print to PDF
  - PDF (.pdf) - Requires weasyprint
  - DOCX (.docx) - Editable in Word

DEFAULT LAYOUT (A4):
  - 10 labels per row
  - 13 labels per column
  - 130 labels per page

LABEL FORMAT:
  Line 1: Location (region/country)
  Line 2: Location (municipality/locality)
  [empty line]
  Code (specimen ID)
  Date (collection date)
"""
    click.echo(info_text)


def main():
    """Entry point for the CLI."""
    cli()


if __name__ == "__main__":
    main()
