"""
Loading and saving layout configuration.

Layout settings are fiddly to get right and tedious to retype: a printed sheet
that lines up with a particular batch of card stock is worth keeping. This
module persists a LabelConfig as JSON.

It lives apart from ``config.py`` because that module holds the constants
LabelConfig itself imports; reading LabelConfig back into it would be a cycle.
"""

import json
import logging
import os
from pathlib import Path
from typing import Any, Dict, Optional, Union

from .config import DEFAULT_CONFIG_PATH, MAX_FILE_SIZE_BYTES
from .label_generator import LabelConfig

logger = logging.getLogger(__name__)

#: Directory name used under the platform's per-user configuration root.
APP_DIR_NAME = "entomology-labels"


def user_config_path() -> Path:
    """Return the per-user configuration file location for this platform.

    Returns:
        Path to the user-level config file; it need not exist
    """
    if os.name == "nt":
        root = os.environ.get("APPDATA")
        base = Path(root) if root else Path.home() / "AppData" / "Roaming"
    else:
        root = os.environ.get("XDG_CONFIG_HOME")
        base = Path(root) if root else Path.home() / ".config"
    return base / APP_DIR_NAME / DEFAULT_CONFIG_PATH


def load_config_file(path: Union[str, Path]) -> LabelConfig:
    """Load a layout configuration from a JSON file.

    The file is parsed with ``json.load`` and fed through
    ``LabelConfig.from_dict``, which keeps only known fields; the dataclass
    then bounds-checks them. A hostile config therefore cannot inject CSS
    through font_family or request a ten-million-label grid.

    Args:
        path: Path to the JSON configuration file

    Returns:
        The configuration described by the file

    Raises:
        ValueError: If the file is missing, too large, malformed, or the
            settings it contains are out of bounds
    """
    config_path = Path(path).expanduser()

    if not config_path.is_file():
        raise ValueError(f"Configuration file not found: {config_path}")

    size = config_path.stat().st_size
    if size > MAX_FILE_SIZE_BYTES:
        raise ValueError(f"Configuration file is too large ({size} bytes): {config_path}")

    try:
        with config_path.open(encoding="utf-8") as handle:
            data = json.load(handle)
    except json.JSONDecodeError as e:
        raise ValueError(f"Configuration file is not valid JSON ({config_path}): {e}")
    except OSError as e:
        raise ValueError(f"Cannot read configuration file {config_path}: {e}")

    if not isinstance(data, dict):
        raise ValueError(f"Configuration file must contain a JSON object: {config_path}")

    try:
        config = LabelConfig.from_dict(data)
    except (TypeError, ValueError) as e:
        raise ValueError(f"Invalid configuration in {config_path}: {e}")

    logger.info(f"Loaded layout configuration from {config_path}")
    return config


def save_config_file(config: LabelConfig, path: Union[str, Path]) -> Path:
    """Write a layout configuration to a JSON file.

    Args:
        config: The configuration to persist
        path: Destination path; parent directories are created

    Returns:
        The resolved path written to

    Raises:
        ValueError: If the file cannot be written
    """
    config_path = Path(path).expanduser()

    try:
        config_path.parent.mkdir(parents=True, exist_ok=True)
        with config_path.open("w", encoding="utf-8") as handle:
            json.dump(config.to_dict(), handle, indent=2, sort_keys=True)
            handle.write("\n")
    except OSError as e:
        raise ValueError(f"Cannot write configuration file {config_path}: {e}")

    logger.info(f"Saved layout configuration to {config_path}")
    return config_path


def find_config_file(explicit_path: Optional[Union[str, Path]] = None) -> Optional[Path]:
    """Find the configuration file that applies, if any.

    Search order, most specific first: the path given on the command line, a
    file in the working directory, then the per-user file.

    Args:
        explicit_path: A path named explicitly by the caller

    Returns:
        The first configuration file found, or None

    Raises:
        ValueError: If an explicitly named file does not exist, since silently
            ignoring it would apply settings the user did not ask for
    """
    if explicit_path is not None:
        candidate = Path(explicit_path).expanduser()
        if not candidate.is_file():
            raise ValueError(f"Configuration file not found: {candidate}")
        return candidate

    for candidate in (Path.cwd() / DEFAULT_CONFIG_PATH, user_config_path()):
        if candidate.is_file():
            return candidate

    return None


def resolve_config(
    explicit_path: Optional[Union[str, Path]] = None,
    overrides: Optional[Dict[str, Any]] = None,
) -> LabelConfig:
    """Build the configuration in effect for a run.

    Precedence, strongest first: values passed on the command line, then the
    configuration file, then the dataclass defaults.

    Args:
        explicit_path: Configuration file named on the command line
        overrides: Settings given explicitly; None values are ignored so that
            an unset command-line option does not mask the file

    Returns:
        The effective configuration

    Raises:
        ValueError: If the file or the resulting settings are invalid
    """
    config_file = find_config_file(explicit_path)
    base = load_config_file(config_file) if config_file else LabelConfig()

    supplied = {k: v for k, v in (overrides or {}).items() if v is not None}
    if not supplied:
        return base

    merged = base.to_dict()
    merged.update(supplied)
    return LabelConfig.from_dict(merged)
