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

# Padding inside each label, in millimetres. The HTML renderer and the fit
# checker both read this, so that a warning about clipped text cannot
# disagree with the layout that actually clips it.
LABEL_PADDING_MM = 1.0

# Fraction of the usable width at which a line is reported as at risk of
# being clipped. Below 1.0 because the width estimate is approximate.
TEXT_WIDTH_WARN_RATIO = 0.9

# How overlong label text is handled: "wrap" onto another line, "clip" it
# with an ellipsis, or "shrink" the font until it fits.
TEXT_OVERFLOW_MODES = ("wrap", "clip", "shrink")
DEFAULT_TEXT_OVERFLOW = "wrap"

# Collection sites
MAX_SITES = 1000  # Maximum sites a single input file may declare
# Site ids appear in error messages and file keys; keep them plain.
SITE_ID_PATTERN = r"^[A-Za-z0-9_-]{1,32}$"

# Geographic bounds. Coordinates come from photo metadata and hand-typed
# files alike, so both need checking before they reach a museum label.
LATITUDE_MIN = -90.0
LATITUDE_MAX = 90.0
LONGITUDE_MIN = -180.0
LONGITUDE_MAX = 180.0
# Roughly the Dead Sea shore to above Everest, with room to spare.
ELEVATION_M_MIN = -500.0
ELEVATION_M_MAX = 9000.0

# Photo metadata scanning. These are separate from MAX_FILE_SIZE_BYTES on
# purpose: a raw frame is legitimately far larger than any label data file,
# so the two limits guard different things and are tuned differently.
MAX_PHOTO_SIZE_BYTES = 200 * 1024 * 1024  # A raw frame; TG-7 ORFs run 15-25MB
MAX_PHOTOS_PER_SCAN = 5000  # Bound on a single directory walk
MAX_EXIF_STRING_LEN = 200  # Camera-written text reaches a printed label
PHOTO_EXTENSIONS = (
    ".orf",
    ".jpg",
    ".jpeg",
    ".tif",
    ".tiff",
    ".dng",
    ".cr2",
    ".cr3",
    ".nef",
    ".arw",
    ".rw2",
    ".heic",
)

# File Paths
DEFAULT_CONFIG_PATH = "label_config.json"

# Logging Configuration
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
LOG_LEVEL = "INFO"
