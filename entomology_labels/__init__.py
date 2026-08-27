"""
Entomology Labels Generator

A tool for generating professional entomology specimen labels with support for
multiple input formats (Excel, CSV, TXT, DOCX, JSON) and output formats (HTML, PDF, DOCX).
"""

try:  # pragma: no cover - trivial packaging shim
    from importlib.metadata import PackageNotFoundError
    from importlib.metadata import version as _pkg_version

    __version__ = _pkg_version("entomology-labels")
except (ImportError, PackageNotFoundError):  # running from an uninstalled source tree
    __version__ = "0.0.0.dev0"
__author__ = "Entomology Labels Generator Contributors"

from .input_handlers import load_data
from .label_generator import Label, LabelConfig, LabelGenerator, expand_label
from .output_generators import generate_docx, generate_html, generate_pdf

__all__ = [
    "LabelGenerator",
    "Label",
    "LabelConfig",
    "expand_label",
    "load_data",
    "generate_html",
    "generate_pdf",
    "generate_docx",
]
