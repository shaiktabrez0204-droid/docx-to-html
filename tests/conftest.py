import os
import sys

# Make the project packages (core, adapter, output, semantic) importable for
# every test in this directory, matching how the vertical-slice test resolves
# imports (the repository root is on sys.path; there is no installed
# `docx_to_html` package).
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)
