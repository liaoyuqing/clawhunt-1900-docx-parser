"""Parse one DOCX file and print its blocks as JSON.

Usage:
    python tools/parse_docx.py samples/demo-a.docx
    python tools/parse_docx.py samples/demo-a.docx --no-expand-merges
"""

from __future__ import annotations

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.parsers.docx_parser import main  # noqa: E402  (path set above)

if __name__ == "__main__":
    sys.exit(main())
