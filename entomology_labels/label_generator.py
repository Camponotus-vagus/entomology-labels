"""
Core label generator module.

Handles the generation of entomology labels with configurable dimensions and layout.
"""

import logging
import math
import re
from dataclasses import dataclass, replace
from typing import ClassVar, List, Optional, Tuple

from .config import (
    DEFAULT_TEXT_OVERFLOW,
    FONT_FAMILY_PATTERN,
    FONT_SIZE_PT_MAX,
    FONT_SIZE_PT_MIN,
    LABEL_HEIGHT_MM_MAX,
    LABEL_HEIGHT_MM_MIN,
    LABEL_PADDING_MM,
    LABEL_WIDTH_MM_MAX,
    LABEL_WIDTH_MM_MIN,
    LABELS_PER_COLUMN_MAX,
    LABELS_PER_COLUMN_MIN,
    LABELS_PER_ROW_MAX,
    LABELS_PER_ROW_MIN,
    MARGIN_MM_MAX,
    MARGIN_MM_MIN,
    MAX_LABELS_PER_GENERATOR,
    MAX_SEQUENTIAL_LABELS,
    TEXT_OVERFLOW_MODES,
)

logger = logging.getLogger(__name__)


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


@dataclass
class LabelConfig:
    """Configuration for label dimensions and layout.

    Default values are optimized for standard A4 paper in LANDSCAPE orientation
    (297mm x 210mm) with 10 labels per row and 13 labels per column.

    Attributes:
        labels_per_row: Number of labels horizontally per page (default: 10)
        labels_per_column: Number of labels vertically per page (default: 13)
        label_width_mm: Width of each label in millimeters (default: 29.0)
        label_height_mm: Height of each label in millimeters (default: 13.0)
        page_width_mm: Page width in millimeters (default: 297 for A4 landscape)
        page_height_mm: Page height in millimeters (default: 210 for A4 landscape)
        margin_top_mm: Top margin in millimeters (default: 0)
        margin_bottom_mm: Bottom margin in millimeters (default: 0)
        margin_left_mm: Left margin in millimeters (default: 0)
        margin_right_mm: Right margin in millimeters (default: 0)
        font_size_pt: Font size in points (default: 6)
        font_family: Font family name (default: Arial)
        line_spacing: Line spacing multiplier (default: 1.0)
        orientation: Page orientation ('landscape' or 'portrait', default: 'landscape')
        label_padding_mm: Padding inside each label (default: 1.0). At small
            label sizes this is a large fraction of the height, so dense
            layouts reduce it
        spacer_line: Whether to leave a blank line between the locality and
            the code (default: True). Dropping it frees a whole line
        text_overflow: How text too wide for the label is handled --
            'wrap' onto another line, 'clip' it with an ellipsis, or
            'shrink' the font until it fits (default: 'wrap')
    """

    labels_per_row: int = 10
    labels_per_column: int = 13
    label_width_mm: float = 29.0  # Default label width
    label_height_mm: float = 13.0  # Default label height
    page_width_mm: float = 297.0  # A4 landscape width
    page_height_mm: float = 210.0  # A4 landscape height
    margin_top_mm: float = 0.0
    margin_bottom_mm: float = 0.0
    margin_left_mm: float = 0.0
    margin_right_mm: float = 0.0
    font_size_pt: float = 6.0
    font_family: str = "Arial"
    line_spacing: float = 1.0
    orientation: str = "landscape"  # 'landscape' or 'portrait'
    text_overflow: str = DEFAULT_TEXT_OVERFLOW
    label_padding_mm: float = LABEL_PADDING_MM
    spacer_line: bool = True

    def __post_init__(self):
        """Validate configuration values after initialization."""
        self._validate()

    def _validate(self) -> None:
        """Validate all configuration values are within safe bounds.

        Raises:
            ValueError: If any configuration value is out of bounds
        """
        if not (LABELS_PER_ROW_MIN <= self.labels_per_row <= LABELS_PER_ROW_MAX):
            raise ValueError(
                f"labels_per_row must be between {LABELS_PER_ROW_MIN} and "
                f"{LABELS_PER_ROW_MAX}, got {self.labels_per_row}"
            )
        if not (LABELS_PER_COLUMN_MIN <= self.labels_per_column <= LABELS_PER_COLUMN_MAX):
            raise ValueError(
                f"labels_per_column must be between {LABELS_PER_COLUMN_MIN} and "
                f"{LABELS_PER_COLUMN_MAX}, got {self.labels_per_column}"
            )
        if not (LABEL_WIDTH_MM_MIN <= self.label_width_mm <= LABEL_WIDTH_MM_MAX):
            raise ValueError(
                f"label_width_mm must be between {LABEL_WIDTH_MM_MIN} and "
                f"{LABEL_WIDTH_MM_MAX}, got {self.label_width_mm}"
            )
        if not (LABEL_HEIGHT_MM_MIN <= self.label_height_mm <= LABEL_HEIGHT_MM_MAX):
            raise ValueError(
                f"label_height_mm must be between {LABEL_HEIGHT_MM_MIN} and "
                f"{LABEL_HEIGHT_MM_MAX}, got {self.label_height_mm}"
            )
        if not (FONT_SIZE_PT_MIN <= self.font_size_pt <= FONT_SIZE_PT_MAX):
            raise ValueError(
                f"font_size_pt must be between {FONT_SIZE_PT_MIN} and "
                f"{FONT_SIZE_PT_MAX}, got {self.font_size_pt}"
            )
        if not (MARGIN_MM_MIN <= self.margin_top_mm <= MARGIN_MM_MAX):
            raise ValueError(
                f"margin_top_mm must be between {MARGIN_MM_MIN} and "
                f"{MARGIN_MM_MAX}, got {self.margin_top_mm}"
            )
        if not (MARGIN_MM_MIN <= self.margin_bottom_mm <= MARGIN_MM_MAX):
            raise ValueError(
                f"margin_bottom_mm must be between {MARGIN_MM_MIN} and "
                f"{MARGIN_MM_MAX}, got {self.margin_bottom_mm}"
            )
        if not (MARGIN_MM_MIN <= self.margin_left_mm <= MARGIN_MM_MAX):
            raise ValueError(
                f"margin_left_mm must be between {MARGIN_MM_MIN} and "
                f"{MARGIN_MM_MAX}, got {self.margin_left_mm}"
            )
        if not (MARGIN_MM_MIN <= self.margin_right_mm <= MARGIN_MM_MAX):
            raise ValueError(
                f"margin_right_mm must be between {MARGIN_MM_MIN} and "
                f"{MARGIN_MM_MAX}, got {self.margin_right_mm}"
            )
        if not re.match(FONT_FAMILY_PATTERN, self.font_family):
            raise ValueError(
                f"font_family must match {FONT_FAMILY_PATTERN} "
                f"(letters, digits, spaces, underscores and hyphens), "
                f"got '{self.font_family}'"
            )
        if self.orientation.lower() not in ("landscape", "portrait"):
            raise ValueError(
                f"orientation must be 'landscape' or 'portrait', " f"got '{self.orientation}'"
            )

        if not (0.0 <= self.label_padding_mm <= MARGIN_MM_MAX):
            raise ValueError(
                f"label_padding_mm must be between 0 and {MARGIN_MM_MAX}, "
                f"got {self.label_padding_mm}"
            )

        if self.text_overflow.lower() not in TEXT_OVERFLOW_MODES:
            raise ValueError(
                f"text_overflow must be one of {TEXT_OVERFLOW_MODES}, "
                f"got '{self.text_overflow}'"
            )

    @property
    def is_landscape(self) -> bool:
        """Check if orientation is landscape."""
        return self.orientation.lower() == "landscape"

    @property
    def labels_per_page(self) -> int:
        """Total number of labels per page."""
        return self.labels_per_row * self.labels_per_column

    @property
    def label_width_pt(self) -> float:
        """Label width in points (1mm = 2.83465pt)."""
        return self.label_width_mm * 2.83465

    @property
    def label_height_pt(self) -> float:
        """Label height in points."""
        return self.label_height_mm * 2.83465

    def to_dict(self) -> dict:
        """Convert config to dictionary."""
        return {
            "labels_per_row": self.labels_per_row,
            "labels_per_column": self.labels_per_column,
            "label_width_mm": self.label_width_mm,
            "label_height_mm": self.label_height_mm,
            "page_width_mm": self.page_width_mm,
            "page_height_mm": self.page_height_mm,
            "margin_top_mm": self.margin_top_mm,
            "margin_bottom_mm": self.margin_bottom_mm,
            "margin_left_mm": self.margin_left_mm,
            "margin_right_mm": self.margin_right_mm,
            "font_size_pt": self.font_size_pt,
            "font_family": self.font_family,
            "line_spacing": self.line_spacing,
            "orientation": self.orientation,
            "text_overflow": self.text_overflow,
            "label_padding_mm": self.label_padding_mm,
            "spacer_line": self.spacer_line,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "LabelConfig":
        """Create config from dictionary."""
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})


@dataclass
class Label:
    """A single specimen label.

    Fields hold text ready to print, not raw values: coordinates and elevation
    arrive already formatted, so that the decision between 46.1873N and
    46 deg 11' 14" N is made once, where the numbers are, rather than
    separately in each renderer.

    Attributes:
        location_line1: Country and region, e.g. "Norway, Vestland,"
        location_line2: Municipality and locality, e.g. "Bergen, Fløyen"
        code: Specimen code, e.g. "N1"
        date: Collection date, e.g. "20.viii.2026"
        additional_info: Anything else worth recording (optional)
        species: Taxon name, printed on the determination label (optional)
        coordinates: Pre-formatted coordinates, e.g. "60.3965N 5.3531E" (optional)
        elevation: Pre-formatted elevation, e.g. "310 m" (optional)
        collector: Who collected it; printed as "leg. ..." (optional)
        determiner: Who identified it; printed as "det. ..." (optional)
        site: Id of the collection site this label refers to. A reference,
            not printed text; resolved against a site registry (optional)
        render_as: Which label this renders as, "locality" or "determination".
            A display concern, not collection data, so it is left out of
            to_dict() and ignored when deciding whether a label is empty.
    """

    location_line1: str = ""
    location_line2: str = ""
    code: str = ""
    date: str = ""
    additional_info: str = ""
    species: str = ""
    coordinates: str = ""
    elevation: str = ""
    collector: str = ""
    determiner: str = ""
    site: str = ""
    render_as: str = "locality"

    #: Fields that carry collection data, as opposed to rendering state.
    CONTENT_FIELDS: ClassVar[Tuple[str, ...]] = (
        "location_line1",
        "location_line2",
        "code",
        "date",
        "additional_info",
        "species",
        "coordinates",
        "elevation",
        "collector",
        "determiner",
        "site",
    )

    def __post_init__(self):
        """Sanitize every text field after initialization."""
        for name in self.CONTENT_FIELDS:
            setattr(self, name, _sanitize_string(getattr(self, name)))

    def is_empty(self) -> bool:
        """Check if the label has no content."""
        return not any(getattr(self, name).strip() for name in self.CONTENT_FIELDS)

    def to_dict(self) -> dict:
        """Convert label to dictionary."""
        return {name: getattr(self, name) for name in self.CONTENT_FIELDS}

    @classmethod
    def from_dict(cls, data: dict) -> "Label":
        """Create label from dictionary, accepting the usual field aliases."""
        aliases = {
            "location_line1": ("location_line1", "location1"),
            "location_line2": ("location_line2", "location2"),
            "code": ("code", "specimen_code"),
            "date": ("date", "collection_date"),
            "additional_info": ("additional_info", "notes"),
            "species": ("species", "taxon"),
            "coordinates": ("coordinates", "coords"),
            "elevation": ("elevation", "altitude"),
            "collector": ("collector", "leg"),
            "determiner": ("determiner", "det"),
            "site": ("site", "site_id"),
        }
        values = {}
        for field_name, names in aliases.items():
            for name in names:
                if name in data and data[name] is not None:
                    values[field_name] = str(data[name])
                    break
        return cls(**values)


def expand_label(label: Label, count: int) -> List[Label]:
    """Return ``count`` independent copies of ``label``.

    Uses dataclasses.replace so that every field is carried over, including
    any added later. Hand-written copies that enumerate the fields silently
    drop new ones, which is why this lives in exactly one place.

    Args:
        label: The label to duplicate
        count: Number of copies to produce

    Returns:
        A list of independent copies; empty when count is not positive
    """
    return [replace(label) for _ in range(count)]


class LabelGenerator:
    """Generator for entomology labels.

    Handles the organization and pagination of labels according to the configuration.
    """

    def __init__(self, config: Optional[LabelConfig] = None):
        """Initialize the label generator.

        Args:
            config: Label configuration. If None, uses default configuration.
        """
        self.config = config or LabelConfig()
        self.labels: List[Label] = []

    def add_label(self, label: Label) -> None:
        """Add a single label to the generator.

        Raises:
            ValueError: If adding this label would exceed the maximum limit
        """
        if len(self.labels) >= MAX_LABELS_PER_GENERATOR:
            raise ValueError(
                f"Maximum label count ({MAX_LABELS_PER_GENERATOR}) exceeded. "
                "Cannot add more labels."
            )
        self.labels.append(label)

    def add_labels(self, labels: List[Label]) -> None:
        """Add multiple labels to the generator.

        Raises:
            ValueError: If adding these labels would exceed the maximum limit
        """
        if len(self.labels) + len(labels) > MAX_LABELS_PER_GENERATOR:
            raise ValueError(
                f"Maximum label count ({MAX_LABELS_PER_GENERATOR}) exceeded. "
                f"Current: {len(self.labels)}, Adding: {len(labels)}"
            )
        self.labels.extend(labels)

    def clear_labels(self) -> None:
        """Remove all labels from the generator."""
        self.labels.clear()

    @property
    def total_labels(self) -> int:
        """Total number of labels."""
        return len(self.labels)

    @property
    def total_pages(self) -> int:
        """Total number of pages needed."""
        if not self.labels:
            return 0
        return math.ceil(len(self.labels) / self.config.labels_per_page)

    def get_labels_for_page(self, page_number: int) -> List[Label]:
        """Get labels for a specific page (0-indexed).

        Args:
            page_number: Page number (0-indexed)

        Returns:
            List of labels for the specified page
        """
        start_idx = page_number * self.config.labels_per_page
        end_idx = start_idx + self.config.labels_per_page
        return self.labels[start_idx:end_idx]

    def get_labels_grid(self, page_number: int) -> List[List[Optional[Label]]]:
        """Get labels organized as a 2D grid for a specific page.

        Args:
            page_number: Page number (0-indexed)

        Returns:
            2D list of labels organized by rows and columns
        """
        page_labels = self.get_labels_for_page(page_number)
        grid = []

        for row in range(self.config.labels_per_column):
            row_labels = []
            for col in range(self.config.labels_per_row):
                idx = row * self.config.labels_per_row + col
                if idx < len(page_labels):
                    row_labels.append(page_labels[idx])
                else:
                    row_labels.append(None)
            grid.append(row_labels)

        return grid

    def expand_label(self, label: Label, count: int) -> List[Label]:
        """Create multiple copies of a label.

        Args:
            label: The label to duplicate
            count: Number of copies

        Returns:
            List of label copies
        """
        return expand_label(label, count)

    def generate_sequential_labels(
        self,
        location_line1: str,
        location_line2: str,
        code_prefix: str,
        start_number: int,
        end_number: int,
        date: str = "",
        additional_info: str = "",
    ) -> List[Label]:
        """Generate a sequence of labels with incrementing codes.

        Args:
            location_line1: First line of location
            location_line2: Second line of location
            code_prefix: Prefix for the code (e.g., "N" for N1, N2, etc.)
            start_number: Starting number for the sequence
            end_number: Ending number for the sequence (inclusive)
            date: Collection date
            additional_info: Additional information

        Returns:
            List of generated labels

        Raises:
            ValueError: If the range is inverted or exceeds the limit
        """
        if end_number < start_number:
            raise ValueError(
                f"end_number ({end_number}) must be greater than or equal to "
                f"start_number ({start_number})"
            )

        # Validate the range to prevent DoS
        count = end_number - start_number + 1
        if count > MAX_SEQUENTIAL_LABELS:
            raise ValueError(
                f"Cannot generate more than {MAX_SEQUENTIAL_LABELS} sequential labels. "
                f"Requested: {count} (from {start_number} to {end_number})"
            )

        logger.info(f"Generating {count} sequential labels with prefix '{code_prefix}'")

        labels = []
        for i in range(start_number, end_number + 1):
            labels.append(
                Label(
                    location_line1=location_line1,
                    location_line2=location_line2,
                    code=f"{code_prefix}{i}",
                    date=date,
                    additional_info=additional_info,
                )
            )
        return labels
