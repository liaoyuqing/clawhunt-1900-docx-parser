"""Regenerate the fictional DOCX samples used by the parser tests.

Every document built here contains invented content only: no real company
name, no real person, no real business data.

Usage:
    python tools/make_samples.py

Files written:
    samples/demo-a.docx                            Chinese heading styles sample
    samples/demo-b.docx                            English heading styles sample
    samples/edge/empty.docx                        empty paragraphs only
    samples/edge/header-footer-comment-only.docx   header, footer and a comment only
    samples/edge/image-only.docx                   one picture, no text

The script is deterministic: running it twice produces the same content.
"""

from __future__ import annotations

import base64
import io
import pathlib
import sys

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Emu

ROOT = pathlib.Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "samples"
EDGE = SAMPLES / "edge"

# 1x1 transparent PNG, used for the picture-only edge case.
PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk"
    "+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
)

CN_TEXT_1 = "本文档为虚构样例，仅用于验证 DOCX 解析器的标题识别与分块能力，"
CN_TEXT_2 = "不代表任何真实企业、真实人员或真实业务资料。"
CN_COMMENT = "虚构审阅意见：本段仅为样例批注，不应出现在解析结果中。"
CN_HEADER = "虚构样例 · 页眉不应出现在解析结果中"
CN_FOOTER = "虚构样例 · 页脚不应出现在解析结果中"


def _tracked_change(paragraph, deleted: str, inserted: str) -> None:
    """Append a w:del and a w:ins revision to a paragraph.

    python-docx has no API for tracked changes, so both elements are written
    directly into the paragraph XML.  The deleted text lives in w:delText and
    the inserted text in an ordinary run inside w:ins.
    """
    deleted_xml = (
        f"<w:del {nsdecls('w')} w:id='9001' w:author='Sample Reviewer' "
        "w:date=\'2026-01-05T09:00:00Z\'>"
        f"<w:r><w:delText xml:space='preserve'>{deleted}</w:delText></w:r>"
        "</w:del>"
    )
    inserted_xml = (
        f"<w:ins {nsdecls('w')} w:id='9002' w:author='Sample Reviewer' "
        "w:date=\'2026-01-05T09:00:00Z\'>"
        f"<w:r><w:t xml:space='preserve'>{inserted}</w:t></w:r>"
        "</w:ins>"
    )
    paragraph._p.append(parse_xml(deleted_xml))
    paragraph._p.append(parse_xml(inserted_xml))


def _build_table(
    document,
    header,
    rows,
    header_merge,
    vertical_merge,
):
    """Build the demo table with one horizontal merge and one vertical merge.

    header_merge    (row, col_start, col_end, value) spans two columns
    vertical_merge  (row_start, row_end, col, value) spans two rows
    """
    table = document.add_table(rows=len(rows) + 1, cols=len(header))
    table.style = "Table Grid"
    for index, value in enumerate(header):
        table.cell(0, index).text = value
    for row_index, row in enumerate(rows, start=1):
        for col_index, value in enumerate(row):
            table.cell(row_index, col_index).text = value

    row, col_start, col_end, value = header_merge
    merged = table.cell(row, col_start).merge(table.cell(row, col_end))
    merged.text = value

    row_start, row_end, col, value = vertical_merge
    merged = table.cell(row_start, col).merge(table.cell(row_end, col))
    merged.text = value
    return table


def build_demo_a(path: pathlib.Path) -> None:
    """Chinese sample: 标题 1 / 标题 2 styles, body text, lists, merged table."""
    document = Document()
    for level in range(1, 7):
        name = "标题 " + str(level)
        if name not in [s.name for s in document.styles]:
            document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)

    document.add_paragraph("会员空间平台知识库文档规范（虚构样例）", style="标题 1")
    document.add_paragraph("一、文档入库说明", style="标题 2")
    document.add_paragraph(CN_TEXT_1 + CN_TEXT_2)
    document.add_paragraph("正文段落按普通样式输出为 paragraph 块。", style="List Bullet")
    document.add_paragraph("项目符号段落输出为 list 块，并标记 ordered 为 false。", style="List Bullet")
    document.add_paragraph("编号段落同样输出为 list 块，并标记 ordered 为 true。", style="List Number")

    _build_table(
        document,
        ["环节", "责任岗位", "周期安排", "周期安排", "状态"],
        [
            ["资料收集", "运营专员", "2", "第 1 周", "已完成"],
            ["内容整理", "内容专员", "5", "第 2 周", "待排期"],
            ["复核发布", "运营专员", "3", "第 3 周", ""],
        ],
        header_merge=(0, 2, 3, "周期安排"),
        vertical_merge=(2, 3, 4, "待排期"),
    )

    section = document.sections[0]
    section.header.paragraphs[0].text = CN_HEADER
    section.footer.paragraphs[0].text = CN_FOOTER
    document.save(path)


def build_demo_b(path: pathlib.Path) -> None:
    """English sample: Heading 1 / Heading 2, tracked changes, a comment."""
    document = Document()
    document.add_paragraph(
        "Member Space Platform Knowledge Base Spec (Fictional Sample)",
        style="Heading 1",
    )
    document.add_paragraph("1. Document Ingestion Notes", style="Heading 2")
    paragraph = document.add_paragraph(
        "This document is a fictional sample used to verify heading detection "
        "and block splitting in the DOCX parser. "
    )
    _tracked_change(
        paragraph,
        deleted="It does not represent any real business material.",
        inserted="It does not describe any real business material.",
    )
    document.add_paragraph("Body paragraphs are emitted as paragraph blocks.", style="List Bullet")
    document.add_paragraph(
        "Bullet paragraphs are emitted as list blocks with ordered set to false.",
        style="List Bullet",
    )
    document.add_paragraph(
        "Numbered paragraphs are emitted as list blocks with ordered set to true.",
        style="List Number",
    )

    _build_table(
        document,
        ["Stage", "Owner", "Schedule", "Schedule", "Status"],
        [
            ["Intake", "Operations", "2", "Week 1", "Done"],
            ["Drafting", "Content", "5", "Week 2", "Scheduled"],
            ["Review", "Operations", "3", "Week 3", ""],
        ],
        header_merge=(0, 2, 3, "Schedule"),
        vertical_merge=(2, 3, 4, "Scheduled"),
    )

    commented = document.paragraphs[2]
    document.add_comment(
        runs=commented.runs,
        text=CN_COMMENT,
        author="Sample Reviewer",
        initials="SR",
    )

    section = document.sections[0]
    section.header.paragraphs[0].text = "Fictional sample - header text must not be parsed"
    section.footer.paragraphs[0].text = "Fictional sample - footer text must not be parsed"
    document.save(path)


def build_empty(path: pathlib.Path) -> None:
    """Edge case: a document that only contains empty paragraphs."""
    document = Document()
    document.add_paragraph("")
    document.add_paragraph("")
    document.add_paragraph("")
    document.save(path)


def build_header_footer_comment_only(path: pathlib.Path) -> None:
    """Edge case: header, footer and a comment, but no body text."""
    document = Document()
    paragraph = document.add_paragraph(" ")
    document.add_comment(
        runs=paragraph.runs,
        text="虚构审阅意见：正文为空，仅用于验证批注被忽略。",
        author="Sample Reviewer",
        initials="SR",
    )
    section = document.sections[0]
    section.header.paragraphs[0].text = "虚构样例 · 只有页眉"
    section.footer.paragraphs[0].text = "虚构样例 · 只有页脚"
    document.save(path)


def build_image_only(path: pathlib.Path) -> None:
    """Edge case: a single picture and no text at all."""
    document = Document()
    document.add_picture(io.BytesIO(PNG_1PX), width=Emu(9525 * 40))
    document.save(path)


def main() -> int:
    EDGE.mkdir(parents=True, exist_ok=True)
    targets = [
        (build_demo_a, SAMPLES / "demo-a.docx"),
        (build_demo_b, SAMPLES / "demo-b.docx"),
        (build_empty, EDGE / "empty.docx"),
        (build_header_footer_comment_only, EDGE / "header-footer-comment-only.docx"),
        (build_image_only, EDGE / "image-only.docx"),
    ]
    for builder, path in targets:
        builder(path)
        print("wrote " + str(path.relative_to(ROOT)).replace("\\", "/"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
