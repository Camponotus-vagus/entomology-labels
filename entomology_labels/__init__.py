"""
Entomology Labels Generator

A tool for generating professional entomology specimen labels with support for
multiple input formats (Excel, CSV, TXT, DOCX, JSON) and output formats (HTML, PDF, DOCX).
"""

__version__ = "1.0.0"
__author__ = "Entomology Labels Generator Contributors"

from .input_handlers import load_data
from .label_generator import Label, LabelConfig, LabelGenerator
from .output_generators import generate_docx, generate_html, generate_pdf

__all__ = [
    "LabelGenerator",
    "Label",
    "LabelConfig",
    "load_data",
    "generate_html",
    "generate_pdf",
    "generate_docx",
]
