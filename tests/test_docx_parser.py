"""Tests for the DOCX parser.

Three groups map onto the three acceptance criteria of the work order:

  AC1  heading_cn     Chinese 标题 1 / 标题 2 styles produce two heading levels
  AC2  merged_table   a table with merged cells becomes a full 2-D array
  AC3  edge           an empty document returns no blocks and does not raise

Run one group with:
    python -m pytest tests/test_docx_parser.py -k heading_cn
    python -m pytest tests/test_docx_parser.py -k merged_table
    python -m pytest tests/test_docx_parser.py -k edge
"""

from __future__ import annotations

import pathlib
import sys

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from app.parsers import DocxParser, ParserRegistry, build_default_registry
from app.parsers.blocks import block_type_sequence
from app.parsers import docx_parser as parser_module
from app.parsers.docx_parser import heading_level, list_kind_from_style, parse_docx


# ---------------------------------------------------------------- helpers --


def _make_doc(tmp_path: pathlib.Path, name: str, builder) -> pathlib.Path:
    path = tmp_path / name
    document = Document()
    builder(document)
    document.save(path)
    return path


def _heading_doc(level_style):
    """Builder: two headings with texts H1/H2 plus one body paragraph."""

    def build(document: Document) -> None:
        for level in range(1, 7):
            style_name = level_style(level)
            if style_name not in [s.name for s in document.styles]:
                document.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
        document.add_paragraph("First Level Title", style=level_style(1))
        document.add_paragraph("Second Level Title", style=level_style(2))
        document.add_paragraph("Body text that is not a heading.")

    return build


def _cn_style(level: int) -> str:
    return "标题 " + str(level)


def _en_style(level: int) -> str:
    return "Heading " + str(level)


def _shape(blocks):
    """(type, level) pairs, the part of a block that must match across styles."""
    return [(b["type"], b.get("level")) for b in blocks]


# ------------------------------------------- AC1: Chinese heading styles --


def test_heading_cn_two_levels(tmp_path):
    """标题 1 and 标题 2 must become headings of level 1 and 2."""
    path = _make_doc(tmp_path, "cn-headings.docx", _heading_doc(_cn_style))
    blocks = parse_docx(path)
    headings = [b for b in blocks if b["type"] == "heading"]
    levels = [b["level"] for b in headings]

    assert levels == [1, 2], "expected exactly two levels 1 and 2, got " + str(levels)
    assert set(levels) == {1, 2}
    print("AC1 chinese heading levels:", levels)


def test_heading_cn_text_is_verbatim(tmp_path):
    """The heading text is copied character for character, not trimmed away."""
    path = _make_doc(tmp_path, "cn-text.docx", _heading_doc(_cn_style))
    texts = [b["text"] for b in parse_docx(path) if b["type"] == "heading"]
    assert texts == ["First Level Title", "Second Level Title"]


def test_heading_cn_not_duplicated_as_paragraph(tmp_path):
    """A heading paragraph must not also appear as a paragraph block."""
    path = _make_doc(tmp_path, "cn-dup.docx", _heading_doc(_cn_style))
    blocks = parse_docx(path)
    heading_texts = {b["text"] for b in blocks if b["type"] == "heading"}
    paragraph_texts = {b["text"] for b in blocks if b["type"] == "paragraph"}
    assert heading_texts & paragraph_texts == set()
    assert "First Level Title" not in paragraph_texts


def test_heading_cn_matches_english_family(tmp_path):
    """Chinese 标题 N and English Heading N are the same family."""
    cn_path = _make_doc(tmp_path, "family-cn.docx", _heading_doc(_cn_style))
    en_path = _make_doc(tmp_path, "family-en.docx", _heading_doc(_en_style))

    cn_blocks = parse_docx(cn_path)
    en_blocks = parse_docx(en_path)

    assert block_type_sequence(cn_blocks) == block_type_sequence(en_blocks)
    assert _shape(cn_blocks) == _shape(en_blocks)
    assert [b["text"] for b in cn_blocks] == [b["text"] for b in en_blocks]
    print("AC1 type sequence:", block_type_sequence(cn_blocks))


def test_heading_cn_sample_snapshot(samples, snapshot):
    """The Chinese sample parses exactly to samples/expected/demo-a.json."""
    blocks = parse_docx(samples / "demo-a.docx")
    assert blocks == snapshot("demo-a")
    print("AC1 snapshot demo-a:", len(blocks), "blocks")


def test_heading_cn_family_disabled_degrades_to_paragraph(tmp_path, monkeypatch):
    """Reverse check: without the 标题 whitelist the headings degrade to text."""
    path = _make_doc(tmp_path, "cn-disabled.docx", _heading_doc(_cn_style))

    original = parser_module.heading_level

    def english_only(style_names, *, allow_chinese=True):
        return original(style_names, allow_chinese=False)

    monkeypatch.setattr(parser_module, "heading_level", english_only)
    blocks = parse_docx(path)
    assert [b["type"] for b in blocks] == ["paragraph", "paragraph", "paragraph"]

    monkeypatch.setattr(parser_module, "heading_level", original)
    restored = parse_docx(path)
    assert [b["type"] for b in restored] == ["heading", "heading", "paragraph"]


def test_heading_level_parser_enumerates_the_whole_family():
    """Both spellings, all six levels, and the negative cases."""
    for level in range(1, 7):
        assert heading_level(["Heading " + str(level)]) == level
        assert heading_level(["标题 " + str(level)]) == level
        assert heading_level(["heading" + str(level)]) == level
    assert heading_level(["Normal"]) is None
    assert heading_level(["Heading 7"]) is None
    assert heading_level(["标题 7"]) is None
    assert heading_level([]) is None
    assert heading_level(["标题 1"], allow_chinese=False) is None


def test_heading_en_sample_snapshot(samples, snapshot):
    """The English sample parses exactly to samples/expected/demo-b.json."""
    blocks = parse_docx(samples / "demo-b.docx")
    assert blocks == snapshot("demo-b")
    print("AC1 snapshot demo-b:", len(blocks), "blocks")


# ------------------------------------------------ AC2: merged cell table --


def test_merged_table_from_sample_is_rectangular(samples):
    """demo-a carries one horizontal and one vertical merge."""
    blocks = parse_docx(samples / "demo-a.docx")
    tables = [b for b in blocks if b["type"] == "table"]
    assert len(tables) == 1

    table = tables[0]
    cells = table["cells"]
    widths = {len(row) for row in cells}
    assert widths == {table["cols"]}, "every row must have the header width"
    assert table["rows"] == len(cells)
    print("AC2 table shape:", table["rows"], "rows x", table["cols"], "cols")


def test_merged_table_horizontal_merge_fills_every_cell(samples):
    """The horizontally merged header value repeats over the whole span."""
    table = [b for b in parse_docx(samples / "demo-a.docx") if b["type"] == "table"][0]
    header = table["cells"][0]
    assert header == ["环节", "责任岗位", "周期安排", "周期安排", "状态"]
    assert header.count("周期安排") == 2


def test_merged_table_vertical_merge_inherits_value(samples):
    """Rows below a vertically merged cell repeat the value above them."""
    table = [b for b in parse_docx(samples / "demo-a.docx") if b["type"] == "table"][0]
    cells = table["cells"]
    assert cells[2][4] == "待排期"
    assert cells[3][4] == cells[2][4], "vertically merged cell must inherit"


def test_merged_table_meta_shape(samples):
    """The table block reports its 2-D shape in rows and cols."""
    table = [b for b in parse_docx(samples / "demo-a.docx") if b["type"] == "table"][0]
    assert table["rows"] == 4
    assert table["cols"] == 5
    assert table["meta"] == {"rows": 4, "cols": 5}


def test_merged_table_is_two_dimensional_array(samples):
    """Every cell is a string and the array is a list of lists."""
    table = [b for b in parse_docx(samples / "demo-b.docx") if b["type"] == "table"][0]
    assert isinstance(table["cells"], list)
    assert all(isinstance(row, list) for row in table["cells"])
    assert all(isinstance(cell, str) for row in table["cells"] for cell in row)
    assert {len(row) for row in table["cells"]} == {table["cols"]}


def test_merged_table_english_sample_shape(samples):
    """The English sample yields the same 4x5 shape as the Chinese one."""
    table = [b for b in parse_docx(samples / "demo-b.docx") if b["type"] == "table"][0]
    assert table["rows"] == 4
    assert table["cols"] == 5
    assert table["cells"][0] == ["Stage", "Owner", "Schedule", "Schedule", "Status"]
    assert table["cells"][2][4] == "Scheduled"
    assert table["cells"][3][4] == "Scheduled"


def test_merged_table_expansion_disabled_breaks_the_grid(samples):
    """Reverse check: without merge expansion the rows are ragged."""
    expanded = [b for b in parse_docx(samples / "demo-a.docx") if b["type"] == "table"][0]
    blocks = DocxParser().parse(samples / "demo-a.docx", expand_merges=False)
    collapsed = [b for b in blocks if b["type"] == "table"][0]

    widths = {len(row) for row in collapsed["cells"]}
    assert len(widths) > 1, "without expansion the row widths must differ"
    assert expanded["cells"][0].count("周期安排") == 2
    assert collapsed["cells"][0].count("周期安排") == 1
    assert expanded["cells"][3][4] == "待排期"
    assert collapsed["cells"][3][4] == ""


# ------------------------------------------------- AC3: empty and ignored --


def test_edge_empty_document_returns_no_blocks(edge):
    """A document with only empty paragraphs produces zero blocks."""
    blocks = parse_docx(edge / "empty.docx")
    assert blocks == []
    print("AC3 empty document blocks:", len(blocks))


def test_edge_empty_document_does_not_raise(edge):
    """Parsing an empty document must not throw."""
    try:
        blocks = parse_docx(edge / "empty.docx")
    except Exception as error:  # noqa: BLE001 - the point of the test
        pytest.fail("parsing an empty document raised " + repr(error))
    assert len(blocks) == 0


def test_edge_header_footer_and_comment_are_ignored(edge):
    """Header, footer and comment text never show up in the blocks."""
    blocks = parse_docx(edge / "header-footer-comment-only.docx")
    assert blocks == []
    joined = " ".join(b["text"] for b in blocks)
    assert "页眉" not in joined
    assert "页脚" not in joined
    assert "批注" not in joined
    print("AC3 header/footer/comment only blocks:", len(blocks))


def test_edge_image_only_document_returns_no_blocks(edge):
    """A picture without text yields no blocks and no error."""
    blocks = parse_docx(edge / "image-only.docx")
    assert blocks == []


def test_edge_empty_guard_is_what_keeps_blocks_empty(edge, monkeypatch):
    """Reverse check: with the empty-text guard removed blocks would appear."""
    parser = DocxParser()
    original = parser._paragraph_block

    def emit_everything(document, paragraph, position):
        block = original(document, paragraph, position)
        if block is None:
            return {"type": "paragraph", "position": position, "text": "", "meta": {}}
        return block

    monkeypatch.setattr(parser, "_paragraph_block", emit_everything)
    blocks = parser.parse(edge / "empty.docx")
    assert len(blocks) == 3, "the guard is what removes the three empty paragraphs"


def test_edge_positions_are_consecutive(samples):
    """position starts at 0 and increases by one per emitted block."""
    blocks = parse_docx(samples / "demo-a.docx")
    positions = [b["position"] for b in blocks]
    assert positions == list(range(len(blocks)))
    assert positions[0] == 0


# ------------------------------------------------------ remaining coverage --


def test_tracked_changes_use_accepted_text(samples):
    """w:del text is dropped and w:ins text is kept."""
    blocks = parse_docx(samples / "demo-b.docx")
    body = [b for b in blocks if b["type"] == "paragraph"][0]["text"]
    assert "It does not describe any real business material." in body
    assert "It does not represent any real business material." not in body
    print("AC4 accepted revision text length:", len(body))


def test_tracked_changes_in_a_generated_document(tmp_path):
    """The same rule on a document built inside the test."""

    def build(document: Document) -> None:
        paragraph = document.add_paragraph("Clause one stays. ")
        paragraph._p.append(
            parse_xml(
                "<w:del " + nsdecls("w") + " w:id='1' w:author='T' w:date='2026-01-01T00:00:00Z'>"
                "<w:r><w:delText>Clause removed.</w:delText></w:r></w:del>"
            )
        )
        paragraph._p.append(
            parse_xml(
                "<w:ins " + nsdecls("w") + " w:id='2' w:author='T' w:date='2026-01-01T00:00:00Z'>"
                "<w:r><w:t>Clause inserted.</w:t></w:r></w:ins>"
            )
        )

    path = _make_doc(tmp_path, "revisions.docx", build)
    text = parse_docx(path)[0]["text"]
    assert "Clause removed." not in text
    assert "Clause inserted." in text
    assert text.startswith("Clause one stays.")


def test_lists_are_split_by_kind(samples):
    """Bullet paragraphs are list blocks with ordered false, numbered true."""
    lists = [b for b in parse_docx(samples / "demo-a.docx") if b["type"] == "list"]
    assert len(lists) == 3
    assert [b["ordered"] for b in lists] == [False, False, True]


def test_list_kind_from_style_enumerates_styles():
    """Style based list detection, English and Chinese names."""
    assert list_kind_from_style(["List Bullet"]) is False
    assert list_kind_from_style(["List Number"]) is True
    assert list_kind_from_style(["项目符号 1"]) is False
    assert list_kind_from_style(["Normal"]) is None


def test_parser_is_registered_in_the_registry():
    """The DOCX parser resolves from the default registry."""
    registry = build_default_registry()
    assert isinstance(registry, ParserRegistry)
    assert registry.mime_types() == [parser_module.MIME_DOCX]
    assert registry.for_mime(parser_module.MIME_DOCX).name == "docx"
    assert registry.for_extension("docx").name == "docx"
    assert registry.for_path("docs/Sample.DOCX").name == "docx"
    assert isinstance(registry.register(DocxParser()), DocxParser)


def test_can_parse_accepts_docx_only():
    """can_parse keys on the extension and the MIME type."""
    parser = DocxParser()
    assert parser.can_parse("a.docx")
    assert parser.can_parse("a.txt", parser_module.MIME_DOCX)
    assert not parser.can_parse("a.pdf")
    assert not parser.can_parse("a.docx", "text/plain")


def test_blocks_are_json_serialisable(samples):
    """Every block is plain data, ready to be stored as a snapshot."""
    import json

    blocks = parse_docx(samples / "demo-a.docx")
    again = json.loads(json.dumps(blocks, ensure_ascii=False))
    assert again == blocks
