"""
Input handlers for various file formats.

Supports: Excel (.xlsx, .xls), CSV, TXT, DOCX, JSON, YAML
"""

import json
import logging
import os
from pathlib import Path
from typing import List, Union

from ..config import MAX_COPIES_PER_ENTRY, MAX_FILE_SIZE_BYTES
from ..label_generator import Label, expand_label

logger = logging.getLogger(__name__)


def _validate_file_path(file_path: Union[str, Path]) -> Path:
    """Check that a path points at a readable file of a workable size.

    This is a usability and resource guard, not a sandbox: the path is
    resolved and accepted wherever it points, which is the correct behaviour
    for a desktop tool whose user picks their own files. It is deliberately
    not confined to a directory and does not reject '..' — callers that need
    confinement must enforce it themselves.

    Args:
        file_path: Path to validate

    Returns:
        Resolved Path object

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If the path is not a file or the file is too large
        PermissionError: If file is not readable
    """
    path = Path(file_path).resolve()

    # Check if file exists
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Check if it's actually a file (not directory)
    if not path.is_file():
        raise ValueError(f"Not a file: {file_path}")

    # Check file size to prevent DoS
    try:
        file_size = path.stat().st_size
        if file_size > MAX_FILE_SIZE_BYTES:
            raise ValueError(
                f"File too large: {file_size / (1024*1024):.2f}MB "
                f"(max: {MAX_FILE_SIZE_BYTES / (1024*1024):.0f}MB)"
            )
    except OSError as e:
        raise ValueError(f"Cannot access file: {e}")

    # Check read permissions
    if not os.access(path, os.R_OK):
        raise PermissionError(f"No read permission: {file_path}")

    return path


def _parse_count(value, default: int = 1) -> int:
    """Parse and bound a user-supplied copy count.

    The count comes straight from an input file, so it is validated here —
    before any list of that size is materialised — rather than relying on the
    MAX_LABELS_PER_GENERATOR check, which only runs once the loader has
    already built (and paid for) the full list.

    Args:
        value: Raw count value from the input file
        default: Value to use when the count is missing

    Returns:
        Count as an int between 0 and MAX_COPIES_PER_ENTRY

    Raises:
        ValueError: If the count is not numeric, negative, or above the limit
    """
    if value is None or value == "":
        return default

    try:
        count = int(float(str(value).strip()))
    except (TypeError, ValueError):
        raise ValueError(f"Invalid count value: {value!r} (expected a number)")

    if count < 0:
        raise ValueError(f"Count cannot be negative: {count}")

    if count > MAX_COPIES_PER_ENTRY:
        raise ValueError(
            f"Count {count} exceeds the maximum of {MAX_COPIES_PER_ENTRY} copies per entry"
        )

    return count


def _sanitize_string(value: str) -> str:
    """Sanitize string input by removing dangerous characters.

    Args:
        value: Input string to sanitize

    Returns:
        Sanitized string with null bytes and control characters removed
    """
    if not value:
        return ""

    # Remove null bytes and control characters except newlines, tabs, carriage returns
    return "".join(c for c in str(value) if c in "\n\r\t" or (ord(c) >= 32 and ord(c) != 127))


def load_data(file_path: Union[str, Path]) -> List[Label]:
    """Load label data from a file.

    Automatically detects the file format based on extension and uses
    the appropriate handler.

    Args:
        file_path: Path to the input file

    Returns:
        List of Label objects

    Raises:
        ValueError: If the file format is not supported
        FileNotFoundError: If the file does not exist
    """
    # Validate and secure the file path
    path = _validate_file_path(file_path)
    logger.info(f"Loading data from: {path}")

    extension = path.suffix.lower()

    handlers = {
        ".xlsx": load_excel,
        ".xls": load_excel,
        ".csv": load_csv,
        ".txt": load_txt,
        ".docx": load_docx,
        ".json": load_json,
        ".yaml": load_yaml,
        ".yml": load_yaml,
    }

    handler = handlers.get(extension)
    if handler is None:
        supported = ", ".join(handlers.keys())
        raise ValueError(
            f"Unsupported file format: {extension}. " f"Supported formats: {supported}"
        )

    try:
        labels = handler(path)
        logger.info(f"Loaded {len(labels)} labels from {path}")
        return labels
    except Exception as e:
        logger.error(f"Error loading data from {path}: {e}", exc_info=True)
        raise


def load_excel(file_path: Path) -> List[Label]:
    """Load labels from an Excel file (.xlsx, .xls).

    Column headings are matched case-insensitively, and several names are
    accepted per field. Where a file carries more than one of them, the first
    listed wins:

    - location_line1: location_line1, location1, location, loc1
    - location_line2: location_line2, location2, loc2
    - code: code, specimen_code, specimen_id, id
    - date: date, collection_date
    - additional_info: additional_info, notes, info (optional)
    - count: count, quantity, copies, n (optional, duplicates the label)

    The Italian headings localita1, localita2, codice, data, data_raccolta,
    note and quantita are also accepted, for spreadsheets written before the
    project moved to English.
    """
    try:
        import pandas as pd
    except ImportError:
        raise ImportError(
            "pandas and openpyxl are required for Excel support. "
            "Install with: pip install pandas openpyxl"
        )

    df = pd.read_excel(file_path)
    return _dataframe_to_labels(df)


def load_csv(file_path: Path) -> List[Label]:
    """Load labels from a CSV file.

    Same column expectations as Excel files.
    Supports both comma and semicolon delimiters.
    """
    try:
        import pandas as pd
    except ImportError:
        raise ImportError("pandas is required for CSV support. " "Install with: pip install pandas")

    # Parse as comma-separated, and only retry with semicolons when that
    # succeeds but collapses into a single column. A genuine parse error is
    # left to propagate, so the reported message describes the real problem
    # rather than a second failure on the wrong delimiter.
    df = pd.read_csv(file_path, delimiter=",")
    if len(df.columns) == 1:
        df = pd.read_csv(file_path, delimiter=";")

    return _dataframe_to_labels(df)


def load_txt(file_path: Path) -> List[Label]:
    """Load labels from a TXT file.

    Supports multiple formats:
    1. Tab-separated values (TSV)
    2. Key-value pairs per label (separated by blank lines)
    3. Simple format: each group of 4-5 lines is a label

    Format 2 example:
        location1: Italia, Trentino Alto Adige,
        location2: Giustino (TN), Vedretta d'Amola
        code: N1
        date: 15.vi.2024

        location1: Italia, Trentino Alto Adige,
        location2: Giustino (TN), Vedretta d'Amola
        code: N2
        date: 15.vi.2024
    """
    content = file_path.read_text(encoding="utf-8")
    lines = content.strip().split("\n")

    # Detect format
    if "\t" in lines[0] and ":" not in lines[0]:
        # TSV format
        try:
            import pandas as pd

            df = pd.read_csv(file_path, delimiter="\t")
            return _dataframe_to_labels(df)
        except Exception:
            pass

    # Check for key-value format
    if ":" in content:
        return _parse_key_value_txt(content)

    # Simple line-based format
    return _parse_simple_txt(lines)


def _parse_key_value_txt(content: str) -> List[Label]:
    """Parse key-value formatted text."""
    labels = []
    blocks = content.strip().split("\n\n")

    for block in blocks:
        if not block.strip():
            continue

        data = {}
        for line in block.strip().split("\n"):
            if ":" in line:
                key, value = line.split(":", 1)
                data[key.strip().lower()] = value.strip()

        if data:
            label = Label(
                location_line1=data.get(
                    "location1", data.get("location_line1", data.get("località1", ""))
                ),
                location_line2=data.get(
                    "location2", data.get("location_line2", data.get("località2", ""))
                ),
                code=data.get("code", data.get("codice", data.get("specimen_code", ""))),
                date=data.get("date", data.get("data", data.get("collection_date", ""))),
                additional_info=data.get(
                    "additional_info", data.get("notes", data.get("note", ""))
                ),
            )

            # Handle count/quantity for duplicates
            count = _parse_count(data.get("count", data.get("quantity", data.get("quantità", 1))))
            labels.extend(expand_label(label, count))

    return labels


def _parse_simple_txt(lines: List[str]) -> List[Label]:
    """Parse simple line-based format (4-5 lines per label)."""
    labels = []
    current_lines = []

    for line in lines:
        line = line.strip()
        if not line:
            if current_lines:
                label = _lines_to_label(current_lines)
                if not label.is_empty():
                    labels.append(label)
                current_lines = []
        else:
            current_lines.append(line)

    # Don't forget the last label
    if current_lines:
        label = _lines_to_label(current_lines)
        if not label.is_empty():
            labels.append(label)

    return labels


def _lines_to_label(lines: List[str]) -> Label:
    """Convert a list of lines to a Label."""
    return Label(
        location_line1=lines[0] if len(lines) > 0 else "",
        location_line2=lines[1] if len(lines) > 1 else "",
        code=lines[2] if len(lines) > 2 else "",
        date=lines[3] if len(lines) > 3 else "",
        additional_info=lines[4] if len(lines) > 4 else "",
    )


def load_docx(file_path: Path) -> List[Label]:
    """Load labels from a Word document (.docx).

    Supports two formats:
    1. Table format: Each row is a label with columns for each field
    2. Paragraph format: Labels separated by blank paragraphs
    """
    try:
        from docx import Document
    except ImportError:
        raise ImportError(
            "python-docx is required for DOCX support. " "Install with: pip install python-docx"
        )

    doc = Document(file_path)
    labels = []

    # Try table format first
    for table in doc.tables:
        if not table.rows:
            continue
        headers = [cell.text.strip().lower() for cell in table.rows[0].cells]

        for row in table.rows[1:]:
            data = {
                headers[i]: cell.text.strip()
                for i, cell in enumerate(row.cells)
                if i < len(headers)
            }
            label = Label.from_dict(data)
            if not label.is_empty():
                labels.append(label)

    # If no tables, try paragraph format
    if not labels:
        paragraphs = [p.text.strip() for p in doc.paragraphs]
        current_lines = []

        for para in paragraphs:
            if not para:
                if current_lines:
                    label = _lines_to_label(current_lines)
                    if not label.is_empty():
                        labels.append(label)
                    current_lines = []
            else:
                current_lines.append(para)

        if current_lines:
            label = _lines_to_label(current_lines)
            if not label.is_empty():
                labels.append(label)

    return labels


def load_json(file_path: Path) -> List[Label]:
    """Load labels from a JSON file.

    Expected format:
    {
        "labels": [
            {
                "location_line1": "Italia, Trentino Alto Adige,",
                "location_line2": "Giustino (TN), Vedretta d'Amola",
                "code": "N1",
                "date": "15.vi.2024",
                "count": 5  // optional, creates duplicates
            },
            ...
        ]
    }

    Or simply an array:
    [
        {"location_line1": "...", ...},
        ...
    ]
    """
    content = file_path.read_text(encoding="utf-8")
    data = json.loads(content)

    if isinstance(data, dict):
        items = data.get("labels", data.get("data", []))
    else:
        items = data

    labels = []
    for item in items:
        label = Label.from_dict(item)
        count = _parse_count(item.get("count", item.get("quantity", 1)))
        labels.extend(expand_label(label, count))

    return labels


def load_yaml(file_path: Path) -> List[Label]:
    """Load labels from a YAML file.

    Same structure as JSON format.
    """
    try:
        import yaml
    except ImportError:
        raise ImportError(
            "PyYAML is required for YAML support. " "Install with: pip install pyyaml"
        )

    content = file_path.read_text(encoding="utf-8")
    data = yaml.safe_load(content)

    if isinstance(data, dict):
        items = data.get("labels", data.get("data", []))
    else:
        items = data

    labels = []
    for item in items:
        label = Label.from_dict(item)
        count = _parse_count(item.get("count", item.get("quantity", 1)))
        labels.extend(expand_label(label, count))

    return labels


def _dataframe_to_labels(df) -> List[Label]:
    """Convert a pandas DataFrame to a list of Labels."""
    # Normalize column names
    df.columns = [str(c).strip().lower() for c in df.columns]

    # Accepted column names per field, in priority order. English names are
    # listed first so that a file carrying both an English and an Italian
    # heading resolves to the English one; the Italian aliases are kept for
    # compatibility with spreadsheets written before the project moved to
    # English, and removing them would break those files.
    column_map = {
        "location_line1": [
            "location_line1",
            "location1",
            "location",
            "loc1",
            "località1",
            "localita1",
        ],
        "location_line2": [
            "location_line2",
            "location2",
            "loc2",
            "località2",
            "localita2",
        ],
        "code": ["code", "specimen_code", "specimen_id", "id", "codice"],
        "date": ["date", "collection_date", "data_raccolta", "data"],
        "additional_info": ["additional_info", "notes", "info", "note"],
        "count": ["count", "quantity", "copies", "n", "quantità", "quantita"],
    }

    def find_column(possible_names):
        for name in possible_names:
            if name in df.columns:
                return name
        return None

    labels = []
    for _, row in df.iterrows():
        data = {}
        for field, possibilities in column_map.items():
            col = find_column(possibilities)
            if col and col in row:
                value = row[col]
                # Only missing values become empty. Testing truthiness here
                # would also discard a legitimate 0 (a specimen coded "0").
                if value is None or (hasattr(value, "__float__") and str(value) == "nan"):
                    data[field] = ""
                else:
                    data[field] = str(value)

        label = Label(
            location_line1=data.get("location_line1", ""),
            location_line2=data.get("location_line2", ""),
            code=data.get("code", ""),
            date=data.get("date", ""),
            additional_info=data.get("additional_info", ""),
        )

        if not label.is_empty():
            count = _parse_count(data.get("count"))

            labels.extend(expand_label(label, count))

    return labels
