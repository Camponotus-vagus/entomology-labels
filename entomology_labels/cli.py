"""
Command Line Interface for Entomology Labels Generator.

Provides commands for generating labels from various input formats.
"""

import logging
import os
from pathlib import Path

import click

from . import __version__
from .config import DEFAULT_TEXT_OVERFLOW, LOG_FORMAT, LOG_LEVEL, TEXT_OVERFLOW_MODES
from .fit import check_fit
from .input_handlers import load_data
from .label_generator import LabelConfig, LabelGenerator
from .output_generators import generate_docx, generate_html, generate_pdf

# Setup logging
logging.basicConfig(format=LOG_FORMAT, level=getattr(logging, LOG_LEVEL))
logger = logging.getLogger(__name__)

SUPPORTED_OUTPUT_FORMATS = [".html", ".pdf", ".docx"]


def _validate_output_path(output: str) -> Path:
    """Resolve the output path and check it can be written to.

    Checked up front so an unwritable destination is reported before the input
    file is parsed and every label rendered, rather than after.

    Args:
        output: Output path as given on the command line

    Returns:
        Resolved output Path

    Raises:
        click.ClickException: If the format is unsupported or the destination
            is not writable
    """
    output_path = Path(output).resolve()

    if output_path.suffix.lower() not in SUPPORTED_OUTPUT_FORMATS:
        raise click.ClickException(
            f"Unsupported output format: {output_path.suffix}. "
            f"Use {', '.join(SUPPORTED_OUTPUT_FORMATS)}"
        )

    if output_path.is_dir():
        raise click.ClickException(f"Output path is a directory: {output_path}")

    # The generators create missing directories, so check the closest
    # existing ancestor is the one that has to be writable.
    ancestor = output_path.parent
    while not ancestor.exists() and ancestor != ancestor.parent:
        ancestor = ancestor.parent

    if not os.access(ancestor, os.W_OK):
        raise click.ClickException(f"Cannot write to output directory: {ancestor}")

    return output_path


def _write_output(generator, output_path: Path, open_after: bool) -> None:
    """Render the generator to the format implied by the output suffix.

    Args:
        generator: LabelGenerator holding the labels to render
        output_path: Destination file, already validated
        open_after: Whether to open the result when it is written

    Raises:
        click.ClickException: If a needed dependency is missing or the write fails
    """
    output_format = output_path.suffix.lower()
    logger.info(f"Generating {output_format[1:].upper()} output")

    try:
        if output_format == ".html":
            generate_html(generator, output_path, open_in_browser=open_after)
        elif output_format == ".pdf":
            generate_pdf(generator, output_path, open_after=open_after)
        elif output_format == ".docx":
            generate_docx(generator, output_path, open_after=open_after)
    except ImportError as e:
        logger.error(f"Missing dependency: {e}")
        raise click.ClickException(str(e))
    except Exception as e:
        logger.exception("Error generating output")
        raise click.ClickException(f"Error generating output: {e}")


@click.group()
@click.version_option(version=__version__, prog_name="entomology-labels")
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


def _report_fit(generator: LabelGenerator, *, strict: bool = False) -> bool:
    """Print any fit warnings for a generator to stderr.

    Args:
        generator: The generator about to produce output
        strict: Whether warnings should be treated as errors

    Returns:
        True when warnings were found and strict was requested
    """
    warnings = check_fit(generator)
    if not warnings:
        return False

    for warning in warnings:
        click.secho(f"warning: {warning}", err=True, fg="yellow")

    if generator.config.text_overflow == "clip":
        click.secho(
            "note: --text-overflow=clip discards the trimmed text; "
            "'wrap' keeps it on another line",
            err=True,
        )

    return strict


@cli.command()
@click.argument("input_file", type=click.Path(exists=True))
@click.option("-o", "--output", required=True, help="Output file path (.html, .pdf, or .docx)")
@click.option("--rows", default=10, type=int, help="Labels per row (default: 10)")
@click.option("--cols", default=13, type=int, help="Labels per column (default: 13)")
@click.option("--label-width", default=21.0, type=float, help="Label width in mm (default: 21.0)")
@click.option(
    "--label-height", default=22.85, type=float, help="Label height in mm (default: 22.85)"
)
@click.option(
    "--page-width", default=210.0, type=float, help="Page width in mm (default: 210 for A4)"
)
@click.option(
    "--page-height", default=297.0, type=float, help="Page height in mm (default: 297 for A4)"
)
@click.option("--font-size", default=6.0, type=float, help="Font size in points (default: 6)")
@click.option("--font-family", default="Arial", help="Font family (default: Arial)")
@click.option(
    "--text-overflow",
    type=click.Choice(TEXT_OVERFLOW_MODES),
    default=DEFAULT_TEXT_OVERFLOW,
    help=f"Handling for text too wide for a label (default: {DEFAULT_TEXT_OVERFLOW})",
)
@click.option(
    "--strict-fit",
    is_flag=True,
    help="Treat fit warnings as errors instead of printing them",
)
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
    text_overflow: str,
    strict_fit: bool,
    open_after: bool,
    verbose: bool,
):
    """Generate labels from an input file.

    INPUT_FILE: Path to the input file (Excel, CSV, TXT, DOCX, JSON, or YAML)

    Examples:

      entomology-labels generate specimens.xlsx -o labels.html

      entomology-labels generate data.csv -o output.pdf --rows 12 --cols 15

      entomology-labels generate labels.json -o output.docx --open
    """
    if verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    input_path = Path(input_file).resolve()
    output_path = _validate_output_path(output)

    logger.info(f"Processing input file: {input_path}")
    logger.info(f"Output file: {output_path}")

    if verbose:
        click.echo(f"Loading data from: {input_path}")

    # Load data
    try:
        labels = load_data(input_path)
    except FileNotFoundError:
        logger.error(f"File not found: {input_path}")
        raise click.ClickException(f"Input file not found: {input_path}")
    except ValueError as e:
        logger.error(f"Invalid input: {e}")
        raise click.ClickException(f"Invalid input: {e}")
    except Exception as e:
        logger.exception("Unexpected error loading data")
        raise click.ClickException(f"Error loading data: {e}")

    if not labels:
        raise click.ClickException("No labels found in input file")

    if verbose:
        click.echo(f"Loaded {len(labels)} labels")

    logger.info(f"Loaded {len(labels)} labels from {input_path}")

    # Configure generator
    try:
        config = LabelConfig(
            labels_per_row=rows,
            labels_per_column=cols,
            label_width_mm=label_width,
            label_height_mm=label_height,
            page_width_mm=page_width,
            page_height_mm=page_height,
            font_size_pt=font_size,
            font_family=font_family,
            orientation="landscape" if page_width > page_height else "portrait",
            text_overflow=text_overflow,
        )
    except ValueError as e:
        raise click.ClickException(f"Invalid configuration: {e}")

    generator = LabelGenerator(config)
    generator.add_labels(labels)

    if verbose:
        click.echo(f"Configuration: {rows}x{cols} labels per page")
        click.echo(f"Total pages: {generator.total_pages}")

    if _report_fit(generator, strict=strict_fit):
        raise click.ClickException(
            "Layout problems reported above; re-run without --strict-fit to generate anyway"
        )

    _write_output(generator, output_path, open_after)

    click.echo(f"Generated {generator.total_labels} labels on {generator.total_pages} pages")
    click.echo(f"Output saved to: {output_path}")
    logger.info(
        f"Successfully generated {generator.total_labels} labels "
        f"on {generator.total_pages} pages"
    )


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
    open_after: bool,
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
    output_path = _validate_output_path(output)

    logger.info(f"Generating sequential labels: {prefix}{start} to {prefix}{end}")

    try:
        config = LabelConfig(labels_per_row=rows, labels_per_column=cols)
    except ValueError as e:
        raise click.ClickException(f"Invalid configuration: {e}")

    generator = LabelGenerator(config)

    try:
        labels = generator.generate_sequential_labels(
            location_line1=location1,
            location_line2=location2,
            code_prefix=prefix,
            start_number=start,
            end_number=end,
            date=date,
        )
        generator.add_labels(labels)
    except ValueError as e:
        logger.error(f"Error generating sequential labels: {e}")
        raise click.ClickException(str(e))

    logger.info(f"Generated {len(labels)} labels")

    _write_output(generator, output_path, open_after)

    click.echo(f"Generated {len(labels)} sequential labels ({prefix}{start} to {prefix}{end})")
    click.echo(f"Output saved to: {output_path}")
    logger.info(f"Successfully generated {len(labels)} sequential labels")


@cli.command()
def gui():
    """Launch the graphical user interface."""
    try:
        from .gui import main as gui_main

        gui_main()
    except ImportError as e:
        raise click.ClickException(
            f"GUI dependencies not available: {e}\n" "Make sure tkinter is installed."
        )


@cli.command()
@click.argument("output_file", type=click.Path())
@click.option(
    "--format",
    "file_format",
    type=click.Choice(["json", "yaml", "csv", "excel"]),
    default="json",
    help="Template format",
)
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
            "location_line1": "Norway, Vestland,",
            "location_line2": "Bergen, Fl\u00f8yen",
            "code": "N1",
            "date": "20.viii.2026",
            "additional_info": "",
            "count": 1,
        },
        {
            "location_line1": "Italy, Trentino-Alto Adige,",
            "location_line2": "Giustino (TN), Vedretta d'Amola",
            "code": "A1",
            "date": "15.vi.2024",
            "additional_info": "",
            "count": 1,
        },
        {
            "location_line1": "Spain, Andaluc\u00eda,",
            "location_line2": "Granada, Sierra Nevada",
            "code": "G1",
            "date": "02.vii.2025",
            "additional_info": "leg. M. Rossi",
            "count": 3,
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
                raise click.ClickException(
                    "pandas and openpyxl are required. Install with: pip install pandas openpyxl"
                )
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
  - location_line1 (or location1, location, loc1)
  - location_line2 (or location2, loc2)
  - code (or specimen_code, specimen_id, id)
  - date (or collection_date)
  - additional_info (or notes, info) - optional
  - count (or quantity, copies, n) - optional, for duplicating labels

  Italian headings are also accepted for older files:
  localita1, localita2, codice, data, data_raccolta, note, quantita

OUTPUT FORMATS:
  - HTML (.html) - Open in browser, print to PDF
  - PDF (.pdf) - Requires weasyprint
  - DOCX (.docx) - Editable in Word

DEFAULT LAYOUT (A4):
  - 10 labels per row
  - 13 labels per column
  - 130 labels per page

LABEL FORMAT:
  Line 1: Location (country, region)
  Line 2: Location (municipality, locality)
  [empty line]
  Code (specimen ID)
  Date (collection date, e.g. 20.viii.2026)

  Dates use the international entomological convention of a
  Roman-numeral month, which avoids the day/month ambiguity of
  all-numeric dates.
"""
    click.echo(info_text)


def main():
    """Entry point for the CLI."""
    cli()


if __name__ == "__main__":
    main()
