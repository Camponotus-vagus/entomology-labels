# Entomology Labels Generator

[![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Generate professional entomology specimen labels with support for multiple input and output formats.

## Features

- **Graphical User Interface (GUI)**: Intuitive interface to create and manage labels with a real-time visual preview.
- **Command Line Interface (CLI)**: For automation and scripting.
- **Multiple Input Formats**: Excel (.xlsx, .xls), CSV, TXT, Word (.docx), JSON, YAML.
- **Multiple Output Formats**: HTML, PDF, Word (.docx).
- **Configurable Layout**: Fully customizable labels per page, dimensions, margins, and typography.
- **Sequential Generation**: Create series of labels with incremental codes (e.g., N1, N2, N3...).

## Label Format

Each label typically contains:
```
Italy, Trentino Alto Adige,        ← Line 1: Main Location
Giustino (TN), Vedretta d'Amola     ← Line 2: Secondary Location
                                     ← Empty line
N1                                   ← Specimen Code
15.vi.2024                          ← Collection Date
```

## Installation

### Option 1: pip (Recommended)

```bash
# Base installation
pip install entomology-labels

# With full support (all formats)
pip install entomology-labels[all]
```

### Option 2: From Source

```bash
git clone https://github.com/Camponotus-vagus/entomology-labels.git
cd entomology-labels
pip install -e .[all]
```

### Optional Dependencies

| Feature | Packages | Installation |
|---------|-----------|---------------|
| Excel | pandas, openpyxl | `pip install entomology-labels[excel]` |
| Word | python-docx | `pip install entomology-labels[docx]` |
| PDF | weasyprint | `pip install entomology-labels[pdf]` |
| YAML | pyyaml | `pip install entomology-labels[yaml]` |
| All | - | `pip install entomology-labels[all]` |

> **Note**: For PDF generation, weasyprint requires additional system dependencies. See [weasyprint documentation](https://doc.courtbouillon.org/weasyprint/stable/first_steps.html).

## Usage

### Graphical User Interface (GUI)

```bash
entomology-labels-gui
```

Or from Python:

```python
from entomology_labels.gui import main
main()
```

The GUI allows you to:
- Add labels manually or through a guided form.
- Import data from various file formats.
- Edit or duplicate labels in the list.
- Configure layout and dimensions with a visual mockup.
- Export to HTML, PDF, or DOCX.

### Command Line Interface (CLI)

```bash
# Generate labels from Excel to HTML
entomology-labels generate data.xlsx -o labels.html

# Generate labels from CSV to PDF
entomology-labels generate data.csv -o labels.pdf

# Generate labels from JSON to Word
entomology-labels generate data.json -o labels.docx

# With custom layout
entomology-labels generate data.xlsx -o labels.html --rows 12 --cols 15

# Open file after generation
entomology-labels generate data.xlsx -o labels.html --open
```

#### Sequential Generation

```bash
entomology-labels sequence \
  --location1 "Italy, Trentino Alto Adige," \
  --location2 "Giustino (TN), Vedretta d'Amola" \
  --prefix N --start 1 --end 50 \
  --date "15.vi.2024" \
  -o labels.html
```

#### Create Templates

```bash
# Create JSON template
entomology-labels template my_data.json

# Create Excel template
entomology-labels template my_data.xlsx --format excel

# Create CSV template
entomology-labels template my_data.csv --format csv
```

### Python API

```python
from entomology_labels import LabelGenerator, Label, LabelConfig
from entomology_labels import load_data, generate_html, generate_pdf, generate_docx

# Layout configuration
config = LabelConfig(
    labels_per_row=10,
    labels_per_column=13,
    font_size_pt=6,
)

# Create generator
generator = LabelGenerator(config)

# Add labels manually
label = Label(
    location_line1="Italy, Trentino Alto Adige,",
    location_line2="Giustino (TN), Vedretta d'Amola",
    code="N1",
    date="15.vi.2024"
)
generator.add_label(label)

# Or load from file
labels = load_data("data.xlsx")
generator.add_labels(labels)

# Generate output
generate_html(generator, "labels.html", open_in_browser=True)
generate_pdf(generator, "labels.pdf")
generate_docx(generator, "labels.docx")
```

#### Sequential Generation

```python
from entomology_labels import LabelGenerator

generator = LabelGenerator()

# Generate N1 to N50
labels = generator.generate_sequential_labels(
    location_line1="Italy, Trentino Alto Adige,",
    location_line2="Giustino (TN), Vedretta d'Amola",
    code_prefix="N",
    start_number=1,
    end_number=50,
    date="15.vi.2024"
)
generator.add_labels(labels)
```

## Input File Formats

### Excel (.xlsx, .xls)

Create a sheet with columns:

| location_line1 | location_line2 | code | date | count |
|----------------|----------------|------|------|-------|
| Italy, Trentino Alto Adige, | Giustino (TN), Vedretta d'Amola | N1 | 15.vi.2024 | 5 |
| Italy, Lombardia, | Sondrio, Valmalenco | O1 | 20.vii.2024 | 3 |

The `count` column is optional and used to create multiple copies of a label.

### CSV

```csv
location_line1,location_line2,code,date,count
"Italy, Trentino Alto Adige,","Giustino (TN), Vedretta d'Amola",N1,15.vi.2024,5
"Italy, Lombardia,","Sondrio, Valmalenco",O1,20.vii.2024,3
```

### JSON

```json
{
  "labels": [
    {
      "location_line1": "Italy, Trentino Alto Adige,",
      "location_line2": "Giustino (TN), Vedretta d'Amola",
      "code": "N1",
      "date": "15.vi.2024",
      "count": 5
    }
  ]
}
```

### TXT (Key-Value Format)

```
location1: Italy, Trentino Alto Adige,
location2: Giustino (TN), Vedretta d'Amola
code: N1
date: 15.vi.2024
count: 5

location1: Italy, Lombardia,
location2: Sondrio, Valmalenco
code: O1
date: 20.vii.2024
count: 3
```

### Alternative Column Names

The software recognizes several variations for column names:

| Field | Accepted Names |
|-------|----------------|
| location_line1 | location1, location, loc1 |
| location_line2 | location2, loc2 |
| code | specimen_code, id, specimen_id |
| date | collection_date, collection_date |
| count | quantity, n, copies |

## Layout Configuration

### Available Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| labels_per_row | 10 | Labels per row |
| labels_per_column | 13 | Labels per column |
| label_width_mm | 29.0 | Label width (mm) |
| label_height_mm | 13.0 | Label height (mm) |
| page_width_mm | 297.0 | Page width (mm) |
| page_height_mm | 210.0 | Page height (mm) |
| font_size_pt | 6.0 | Font size (pt) |
| font_family | Arial | Font family |

### Presets

- **A4 Standard (Landscape)**: 10x13 labels (130 per page)
- **A4 Compact (Landscape)**: 12x15 labels (180 per page)
- **US Letter (Landscape)**: 10x12 labels (120 per page)

## HTML Output and PDF Printing

HTML output includes a "Print" button that opens the browser's print dialog. To save as PDF:

1. Generate the HTML file.
2. Open it in your browser.
3. Click "Print" or use Ctrl+P (Cmd+P on Mac).
4. Select "Save as PDF" as the destination.
5. Ensure margins are set to "None".

## Examples

The `examples/` directory contains sample files:

- `example_labels.json` - JSON format
- `example_labels.csv` - CSV format
- `example_labels.txt` - TXT format

## Troubleshooting

### "weasyprint not found" Error

```bash
# Ubuntu/Debian
sudo apt-get install libpango-1.0-0 libpangocairo-1.0-0 libgdk-pixbuf2.0-0

# macOS
brew install pango

# Then install weasyprint
pip install weasyprint
```

### "pandas not found" Error

```bash
pip install pandas openpyxl
```

### Labels are not aligned correctly

- Ensure printer margins are set to 0.
- Use "Fit to page" mode if necessary.
- Adjust `margin_*` parameters in the configuration.

## Contributing

Contributions are welcome! To contribute:

1. Fork the repository.
2. Create a feature branch (`git checkout -b feature/NewFeature`).
3. Commit your changes (`git commit -m 'Add NewFeature'`).
4. Push to the branch (`git push origin feature/NewFeature`).
5. Open a Pull Request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Credits

Inspired by [insect-labels](https://github.com/tracyyao27/insect-labels) by Tracy Yao.
