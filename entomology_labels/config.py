"""
Configuration constants and security limits for Entomology Labels Generator.
"""

# Security Limits
MAX_LABELS_PER_GENERATOR = 100000  # Maximum labels in a single generation run
MAX_SEQUENTIAL_LABELS = 10000  # Maximum sequential labels (end - start)
MAX_FILE_SIZE_BYTES = 100 * 1024 * 1024  # 100MB max file size
MAX_DISPLAYED_LABELS = 500  # Maximum labels shown in GUI preview

# Validation Bounds
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
PREVIEW_SCALE_FACTOR = 3.5

# File Paths
DEFAULT_CONFIG_PATH = "label_config.json"

# Logging Configuration
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
LOG_LEVEL = "INFO"
