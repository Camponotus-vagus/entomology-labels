"""
Configuration constants and security limits for Entomology Labels Generator.
"""

# Security Limits
MAX_LABELS_PER_GENERATOR = 100000  # Maximum labels in a single generation run
MAX_SEQUENTIAL_LABELS = 10000  # Maximum sequential labels (end - start)
MAX_FILE_SIZE_BYTES = 100 * 1024 * 1024  # 100MB max file size
MAX_DISPLAYED_LABELS = 500  # Maximum labels shown in GUI preview
MAX_COPIES_PER_ENTRY = 10000  # Maximum copies a single input row may request

# Validation Bounds
LABELS_PER_ROW_MIN = 1
LABELS_PER_ROW_MAX = 100
LABELS_PER_COLUMN_MIN = 1
LABELS_PER_COLUMN_MAX = 100
LABEL_WIDTH_MM_MIN = 1.0
LABEL_WIDTH_MM_MAX = 500.0
LABEL_HEIGHT_MM_MIN = 1.0
LABEL_HEIGHT_MM_MAX = 500.0
FONT_SIZE_PT_MIN = 1.0
FONT_SIZE_PT_MAX = 72.0
MARGIN_MM_MIN = 0.0
MARGIN_MM_MAX = 50.0
DPI_MIN = 72
DPI_MAX = 1200
# font_family is interpolated into a CSS declaration in the generated HTML, so
# it is restricted to characters that cannot terminate the declaration.
FONT_FAMILY_PATTERN = r"^[A-Za-z0-9 _-]+$"
PREVIEW_SCALE_FACTOR = 3.5

# File Paths
DEFAULT_CONFIG_PATH = "label_config.json"

# Logging Configuration
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
LOG_LEVEL = "INFO"
