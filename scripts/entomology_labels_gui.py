#!/usr/bin/env python3
"""Entry point script for PyInstaller - GUI version."""

import sys
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from entomology_labels.gui import main

if __name__ == "__main__":
    main()
