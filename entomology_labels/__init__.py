"""
Entomology Labels Generator

A tool for generating professional entomology specimen labels with support for
multiple input formats (Excel, CSV, TXT, DOCX, JSON) and output formats (HTML, PDF, DOCX).
"""

__version__ = "1.0.0"
__author__ = "Entomology Labels Generator Contributors"

from .label_generator import LabelGenerator, Label, LabelConfig
from .input_handlers import load_data
from .output_generators import generate_html, generate_pdf, generate_docx

__all__ = [
    "LabelGenerator",
    "Label",
    "LabelConfig",
    "load_data",
    "generate_html",
    "generate_pdf",
    "generate_docx",
]
