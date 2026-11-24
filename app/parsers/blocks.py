"""Block model shared by the parsers of the knowledge-base pipeline.

Every parser returns a list of plain dictionaries.  A block always carries:

  type      heading | paragraph | list | table
  position  0-based ordinal of the block inside the document content flow
  text      plain text of the block

Optional keys:

  level     1..6, heading blocks only
  ordered   True (numbered list) / False (bullet list), list blocks only
  rows/cols shape of the 2-D array, table blocks only
  cells     2-D array of strings with merged cells expanded, table blocks only
  meta      small dict with parser side details (style name, table shape)

DOCX carries no reliable page numbers, therefore the position of a block is
the ordinal of the block among the content elements of the document body.
It starts at 0 and increases by 1 for every emitted block, so positions are
always consecutive.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List


@dataclass
class Block:
    type: str
    position: int
    text: str = ""
    level: int | None = None
    ordered: bool | None = None
    rows: int | None = None
    cols: int | None = None
    cells: List[List[str]] = field(default_factory=list)
    meta: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {
            "type": self.type,
            "position": self.position,
            "text": self.text,
        }
        if self.level is not None:
            data["level"] = self.level
        if self.ordered is not None:
            data["ordered"] = self.ordered
        if self.rows is not None:
            data["rows"] = self.rows
        if self.cols is not None:
            data["cols"] = self.cols
        if self.type == "table":
            data["cells"] = self.cells
        data["meta"] = self.meta
        return data


def block_type_sequence(blocks: List[Dict[str, Any]]) -> List[str]:
    """Type sequence of a block list, used by tests and by the CLI."""
    return [b["type"] for b in blocks]
