"""DOCX parser: headings, paragraphs, bullet lists and tables as blocks.

Reading rules implemented here:
  * headings are recognised from the paragraph style; the English family
    Heading 1..6 and the Chinese family 标题 1..6 are one and the same family
  * body paragraphs become paragraph blocks
  * bullet and numbered list paragraphs become list blocks (ordered flag)
  * tables become table blocks, cell by cell, with merged cells expanded:
    a horizontal merge fills every grid cell of the span on that row so that
    every row has as many columns as the header, and a vertically merged cell
    inherits the value of the cell above it
  * headers, footers and comments live in separate parts of the package and
    are never read
  * tracked changes are read in accepted state: inserted runs (w:ins) are kept,
    deleted runs (w:del / w:delText) are dropped
  * DOCX has no reliable page numbers, so position is the ordinal of the block
    inside the document content flow, starting at 0 and always consecutive
  * an empty document returns an empty block list and never raises
"""

from __future__ import annotations

import argparse
import io
import json
import re
import sys
from typing import Any, Dict, Iterable, List, Optional, Sequence

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph

from app.parsers.blocks import Block

MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

# The heading family: English "Heading 1..6" and Chinese "标题 1..6".
_HEADING_EN = re.compile(r"^heading[\s\-\u00a0]*([1-6])$", re.IGNORECASE)
_HEADING_CN = re.compile(r"^标题[\s\-\u00a0]*([1-6])$")

_BULLET_STYLE_HINTS = ("list bullet", "bullet", "项目符号")
_NUMBER_STYLE_HINTS = ("list number", "numbered", "编号")


def heading_level(
    style_names: Sequence[Optional[str]],
    *,
    allow_chinese: bool = True,
) -> Optional[int]:
    """Return 1..6 when one of the style names belongs to the heading family."""
    for raw in style_names:
        if not raw:
            continue
        name = raw.replace("\u3000", " ").strip()
        match = _HEADING_EN.match(name)
        if match:
            return int(match.group(1))
        if allow_chinese:
            match = _HEADING_CN.match(name)
            if match:
                return int(match.group(1))
    return None


def list_kind_from_style(style_names: Sequence[Optional[str]]) -> Optional[bool]:
    """True for numbered lists, False for bullet lists, None when unknown."""
    for raw in style_names:
        if not raw:
            continue
        name = raw.replace("\u3000", " ").strip().lower()
        if not name:
            continue
        if any(hint in name for hint in _NUMBER_STYLE_HINTS):
            return True
        if any(hint in name for hint in _BULLET_STYLE_HINTS):
            return False
    return None


def _collect_text(element) -> str:
    """Plain text of an XML element in accepted-revision state.

    w:del subtrees are skipped entirely (their text lives in w:delText),
    w:ins subtrees are read like normal runs, so inserted text shows up.
    Comments live in their own part and never appear inside the body.
    """
    parts: List[str] = []

    def walk(node) -> None:
        for child in node:
            tag = child.tag
            if tag == qn("w:del"):
                continue
            if tag == qn("w:t"):
                parts.append(child.text or "")
                continue
            if tag == qn("w:tab"):
                parts.append("\t")
                continue
            if tag in (qn("w:br"), qn("w:cr")):
                parts.append("\n")
                continue
            if tag in (qn("w:delText"), qn("w:instrText")):
                continue
            walk(child)

    walk(element)
    return "".join(parts)


def _style_names(paragraph: DocxParagraph) -> List[str]:
    """All candidate names of the style of a paragraph.

    The UI name ("标题 1"), the internal name stored in styles.xml
    ("heading 1") and the style id ("Heading1") are collected, because a
    Chinese Word installation writes the Chinese UI name while the file may
    still carry the internal English name.
    """
    names: List[str] = []
    try:
        style = paragraph.style
    except Exception:
        return names
    if style is None:
        return names
    ui_name = getattr(style, "name", None)
    if ui_name:
        names.append(ui_name)
    try:
        element = style.element
        internal = element.find(qn("w:name"))
        if internal is not None and internal.get(qn("w:val")):
            names.append(internal.get(qn("w:val")))
    except Exception:
        pass
    style_id = getattr(style, "style_id", None)
    if style_id:
        names.append(style_id)
    return names


def _num_id(paragraph: DocxParagraph) -> Optional[str]:
    """Numbering id of a paragraph, when the paragraph uses a list."""
    ppr = paragraph._p.pPr
    if ppr is None:
        return None
    numpr = ppr.find(qn("w:numPr"))
    if numpr is None:
        return None
    num_id = numpr.find(qn("w:numId"))
    if num_id is None:
        return None
    return num_id.get(qn("w:val"))


def _numbering_format(document, num_id: Optional[str]) -> Optional[str]:
    """numFmt of a numbering id ("bullet", "decimal", ...)."""
    if not num_id:
        return None
    try:
        numbering = document.part.numbering_part.element
    except Exception:
        return None
    abstract_id = None
    for num in numbering.findall(qn("w:num")):
        if num.get(qn("w:numId")) == str(num_id):
            link = num.find(qn("w:abstractNumId"))
            if link is not None:
                abstract_id = link.get(qn("w:val"))
    if abstract_id is None:
        return None
    for abstract in numbering.findall(qn("w:abstractNum")):
        if abstract.get(qn("w:abstractNumId")) != abstract_id:
            continue
        level = abstract.find(qn("w:lvl"))
        if level is None:
            return None
        fmt = level.find(qn("w:numFmt"))
        if fmt is None:
            return None
        return fmt.get(qn("w:val"))
    return None


def _paragraph_list_kind(document, paragraph: DocxParagraph, style_names: Sequence[str]) -> Optional[bool]:
    """True for numbered lists, False for bullet lists, None for body text."""
    kind = list_kind_from_style(style_names)
    if kind is not None:
        return kind
    num_id = _num_id(paragraph)
    if num_id is None:
        return None
    fmt = _numbering_format(document, num_id)
    return False if fmt == "bullet" else True


def _cell_text(tc) -> str:
    """Text of a table cell: paragraphs joined by a newline."""
    chunks: List[str] = []
    for paragraph in tc.findall(qn("w:p")):
        text = _collect_text(paragraph).strip()
        if text:
            chunks.append(text)
    return "\n".join(chunks)


def _grid_span(tc) -> int:
    """Width of a cell in grid columns (w:gridSpan), 1 when not merged."""
    tc_pr = tc.find(qn("w:tcPr"))
    if tc_pr is None:
        return 1
    span = tc_pr.find(qn("w:gridSpan"))
    if span is None:
        return 1
    try:
        return max(1, int(span.get(qn("w:val"))))
    except (TypeError, ValueError):
        return 1


def _vmerge(tc) -> Optional[str]:
    """Vertical merge state of a cell: "restart", "continue" or None."""
    tc_pr = tc.find(qn("w:tcPr"))
    if tc_pr is None:
        return None
    merge = tc_pr.find(qn("w:vMerge"))
    if merge is None:
        return None
    value = merge.get(qn("w:val"))
    if value is None or value == "continue":
        return "continue"
    return value


def _table_grid(table: DocxTable, expand_merges: bool = True) -> List[List[str]]:
    """Expand a table into a rectangular 2-D array of strings.

    Horizontal merges (w:gridSpan) repeat the value of the merged cell over
    every column of the span, so each row has as many columns as the header.
    Vertical merges (w:vMerge) make the rows below the merged cell inherit the
    value of the cell above.
    """
    grid: List[List[str]] = []
    for tr in table._tbl.tr_lst:
        row: List[str] = []
        for tc in tr.tc_lst:
            span = _grid_span(tc) if expand_merges else 1
            merge = _vmerge(tc) if expand_merges else None
            value = _cell_text(tc)
            if merge == "continue":
                column = len(row)
                previous = grid[-1] if grid else []
                inherited = previous[column] if column < len(previous) else ""
                if not value:
                    value = inherited
                for offset in range(span):
                    row.append(value)
                    value = inherited
                continue
            for _ in range(span):
                row.append(value)
        grid.append(row)
    if expand_merges:
        width = max((len(row) for row in grid), default=0)
        for row in grid:
            while len(row) < width:
                row.append("")
    return grid


class DocxParser:
    """Parses a .docx file into a list of block dictionaries."""

    name = "docx"
    mime_types = (MIME_DOCX,)
    extensions = ("docx",)
    supports = ("heading", "paragraph", "list", "table")

    def can_parse(self, path: str = "", mime: str = "") -> bool:
        if mime:
            return mime.lower() == MIME_DOCX
        return str(path).lower().endswith(".docx")

    def parse(self, path: str, *, expand_merges: bool = True) -> List[Dict[str, Any]]:
        document = Document(path)
        return self.parse_document(document, expand_merges=expand_merges)

    def parse_bytes(self, data: bytes, *, expand_merges: bool = True) -> List[Dict[str, Any]]:
        document = Document(io.BytesIO(data))
        return self.parse_document(document, expand_merges=expand_merges)

    def parse_document(
        self,
        document,
        *,
        expand_merges: bool = True,
    ) -> List[Dict[str, Any]]:
        """Parse an already opened document into blocks."""
        blocks: List[Dict[str, Any]] = []
        for child in document.element.body.iterchildren():
            if child.tag == qn("w:p"):
                block = self._paragraph_block(document, DocxParagraph(child, document), len(blocks))
            elif child.tag == qn("w:tbl"):
                block = self._table_block(DocxTable(child, document), len(blocks), expand_merges=expand_merges)
            else:
                continue
            if block is not None:
                blocks.append(block)
        return blocks

    def _paragraph_block(self, document, paragraph: DocxParagraph, position: int) -> Optional[Dict[str, Any]]:
        names = _style_names(paragraph)
        text = _collect_text(paragraph._p).strip()
        if not text:
            return None
        level = heading_level(names)
        if level is not None:
            return Block(type="heading", position=position, text=text, level=level, meta={"style": names[0]}).to_dict()
        ordered = _paragraph_list_kind(document, paragraph, names)
        if ordered is not None:
            return Block(type="list", position=position, text=text, ordered=ordered, meta={"style": names[0]}).to_dict()
        return Block(type="paragraph", position=position, text=text, meta={"style": names[0]}).to_dict()

    def _table_block(self, table: DocxTable, position: int, *, expand_merges: bool = True) -> Optional[Dict[str, Any]]:
        cells = _table_grid(table, expand_merges=expand_merges)
        if not cells or not any(any(cell for cell in row) for row in cells):
            return None
        rows = len(cells)
        cols = len(cells[0])
        text = "\n".join("\t".join(row) for row in cells)
        return Block(
            type="table",
            position=position,
            text=text,
            rows=rows,
            cols=cols,
            cells=cells,
            meta={"rows": rows, "cols": cols},
        ).to_dict()


def parse_docx(path: str, *, expand_merges: bool = True) -> List[Dict[str, Any]]:
    """Module level helper: parse one .docx file into blocks."""
    return DocxParser().parse(path, expand_merges=expand_merges)


def main(argv: Optional[List[str]] = None) -> int:
    """Print the blocks of a .docx file as JSON."""
    parser = argparse.ArgumentParser(description="Parse a DOCX file into blocks")
    parser.add_argument("path", help="path of the .docx file")
    parser.add_argument("--no-expand-merges", action="store_true", help="disable merged cell expansion")
    args = parser.parse_args(argv)
    blocks = parse_docx(args.path, expand_merges=not args.no_expand_merges)
    print(json.dumps(blocks, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
