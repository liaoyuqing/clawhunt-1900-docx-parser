"""Regenerate the expected-output snapshots from the two demo samples.

Usage:
    python tools/make_snapshots.py

Writes samples/expected/demo-a.json and samples/expected/demo-b.json.
Run it only after an intentional parser change: the snapshots are the
contract that tests/test_docx_parser.py checks the parser against.
"""

from __future__ import annotations

import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.parsers.docx_parser import parse_docx  # noqa: E402  (path set above)

SAMPLES = ROOT / "samples"
EXPECTED = SAMPLES / "expected"


def main() -> int:
    EXPECTED.mkdir(parents=True, exist_ok=True)
    for name in ("demo-a", "demo-b"):
        blocks = parse_docx(SAMPLES / (name + ".docx"))
        target = EXPECTED / (name + ".json")
        target.write_text(
            json.dumps(blocks, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        print("wrote " + str(target.relative_to(ROOT)).replace("\\", "/") + " (" + str(len(blocks)) + " blocks)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
