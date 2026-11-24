# DOCX document parser (knowledge-base ingestion)

A DOCX parser that turns a Word document into a flat list of content blocks,
plus the parser registry entry the knowledge-base ingestion pipeline looks the
parser up through. Built for the work order **S3-KB-005 · DOCX document
parser**.

All sample documents, company names, people and numbers in this repository
are invented. Nothing here comes from a real organisation.

---

## 1. What the parser does

`app/parsers/docx_parser.py` walks the body of a `.docx` file in document
order and emits blocks of four types:

| Block type  | Produced from | Extra keys |
|-------------|---------------|------------|
| `heading`   | paragraphs whose style is `Heading 1`–`Heading 6` or `标题 1`–`标题 6` | `level` (1–6) |
| `paragraph` | ordinary body paragraphs | – |
| `list`      | bullet and numbered paragraphs (`List Bullet`, `List Number`, or a numbering definition) | `ordered` (`false` = bullet, `true` = numbered) |
| `table`     | tables, cell by cell, merged cells expanded | `rows`, `cols`, `cells` (2-D array of strings) |

Every block also carries `type`, `position` and `text`, and a small `meta`
object (the source style name, and `rows`/`cols` for tables).

Rules that the implementation commits to:

* **One heading family.** `Heading 1`–`Heading 6` and `标题 1`–`标题 6` are the
  same family and map to the same `level`. The style UI name, the internal
  `w:name` in `styles.xml` and the style id are all consulted, so a file
  produced by a Chinese Word installation and a file produced by an English
  one are parsed identically. `app/parsers/docx_parser.py:47` (`heading_level`)
  is the single place that decides this.
* **Merged cells are expanded.** A horizontal merge (`w:gridSpan`) repeats the
  merged value over every column of the span, so every row has as many columns
  as the header. A vertical merge (`w:vMerge`) makes the rows below inherit the
  value of the cell above, even when the lower cell is stored empty.
* **Headers, footers and comments are ignored.** They live in separate parts of
  the OOXML package and are never read, so their text cannot leak into a block.
* **Tracked changes are read in accepted state.** `w:ins` runs are kept,
  `w:del` / `w:delText` runs are dropped. A clause that a reviewer deleted does
  not appear in the output; a clause they inserted does.
* **No page numbers.** DOCX has no reliable page number, so `position` is the
  ordinal of the block in the content flow. **The count starts at 0** and
  increases by exactly 1 per emitted block, so positions are always
  consecutive (`0, 1, 2, …`). Empty paragraphs and unsupported elements are
  skipped and do not consume a position.
* **Empty documents are not an error.** A document with no content — empty
  paragraphs only, header/footer/comment only, or a picture only — returns
  `[]` and does not raise.

---

## 2. Requirements

| Item | Version |
|------|---------|
| Python | 3.11 or newer (tested on 3.11.9) |
| python-docx | 1.2.0 |
| pytest | 9.2.0 |

Nothing else is needed: no database, no network access, no API key, no
environment variable. The parser reads a local file and returns plain
dictionaries. No third-party SDK that sends data anywhere is used.

---

## 3. Run it

```bash
git clone https://github.com/liaoyuqing/clawhunt-1900-docx-parser.git
cd clawhunt-1900-docx-parser
python -m pip install -r requirements.txt
python -m pytest -q
```

Expected tail of the pytest run:

```
28 passed in 0.7s
```

Parse one file from the command line and print its blocks as JSON:

```bash
python tools/parse_docx.py samples/demo-a.docx
```

From Python:

```python
from app.parsers import build_default_registry
from app.parsers.docx_parser import parse_docx

blocks = parse_docx("samples/demo-a.docx")
print([(b["type"], b["position"]) for b in blocks])

registry = build_default_registry()
parser = registry.for_extension("docx")   # -> DocxParser
blocks = parser.parse("samples/demo-b.docx")
```

---

## 4. Acceptance criteria and where they are tested

Test file: `tests/test_docx_parser.py` (one file, 28 test cases).
`tests/conftest.py` only provides fixtures and puts the repository root on
`sys.path`. Logs of the runs quoted below are in `tests/reports/`.

### AC1 — Chinese `标题 1` / `标题 2` styles produce two heading levels

Selection: `python -m pytest tests/test_docx_parser.py -k heading_cn -v`
(log: `tests/reports/ac1-heading.txt`), 6 cases, all passing:

* `test_heading_cn_two_levels` — levels are exactly `[1, 2]`.
* `test_heading_cn_text_is_verbatim` — heading text is character-for-character.
* `test_heading_cn_not_duplicated_as_paragraph` — a heading is not also emitted
  as a paragraph.
* `test_heading_cn_matches_english_family` — a Chinese-styled document and an
  English-styled document produce the same block-type sequence and the same
  `(type, level)` shape; this is the equivalence proof for the shared family.
* `test_heading_cn_sample_snapshot` — `samples/demo-a.docx` matches
  `samples/expected/demo-a.json` exactly.
* `test_heading_cn_family_disabled_degrades_to_paragraph` — reverse check: with
  the Chinese styles removed from the whitelist the same document degrades to
  three `paragraph` blocks, and the test fails if the whitelist is missing.

The English sibling of the snapshot test is
`test_heading_en_sample_snapshot` (`samples/demo-b.docx` against
`samples/expected/demo-b.json`), and
`test_heading_level_parser_enumerates_the_whole_family` covers all twelve
spellings plus the negative cases (`Normal`, `Heading 7`, `标题 7`).

### AC2 — a table with merged cells becomes a full 2-D array

Selection: `python -m pytest tests/test_docx_parser.py -k merged_table -v`
(log: `tests/reports/ac2-table.txt`), 7 cases, all passing:

* `test_merged_table_from_sample_is_rectangular` — every row has `cols` cells.
* `test_merged_table_horizontal_merge_fills_every_cell` — the merged header
  value `周期安排` occupies both columns of its span.
* `test_merged_table_vertical_merge_inherits_value` — the lower row of the
  vertical merge repeats the value above it although the cell is stored empty.
* `test_merged_table_meta_shape` — the block reports `rows: 4`, `cols: 5`.
* `test_merged_table_is_two_dimensional_array` — list of lists of strings.
* `test_merged_table_english_sample_shape` — the English sample has the same
  4×5 shape.
* `test_merged_table_expansion_disabled_breaks_the_grid` — reverse check: with
  merge expansion disabled the rows become ragged (`5` and `4` cells) and the
  inherited value is empty, so the test fails if expansion is removed.

### AC3 — an empty document returns no blocks and does not raise

Selection: `python -m pytest tests/test_docx_parser.py -k edge -v`
(log: `tests/reports/ac3-edge.txt`), 6 cases, all passing:

* `test_edge_empty_document_returns_no_blocks` — `blocks == []`.
* `test_edge_empty_document_does_not_raise` — no exception.
* `test_edge_header_footer_and_comment_are_ignored` — a document with only a
  header, a footer and a comment yields `[]`; none of the three texts appear.
* `test_edge_image_only_document_returns_no_blocks` — a picture-only document
  yields `[]`.
* `test_edge_empty_guard_is_what_keeps_blocks_empty` — reverse check: with the
  empty-text guard removed the same document yields 3 blocks.
* `test_edge_positions_are_consecutive` — `position` is `0..n-1` with no gaps.

### Remaining cases

`test_tracked_changes_use_accepted_text` and
`test_tracked_changes_in_a_generated_document` cover accepted-revision
reading (`w:ins` kept, `w:del` dropped). `test_lists_are_split_by_kind` and
`test_list_kind_from_style_enumerates_styles` cover bullet versus numbered
lists. `test_parser_is_registered_in_the_registry`, `test_can_parse_accepts_docx_only`
and `test_blocks_are_json_serialisable` cover the registry wiring and the data
shape.

---

## 5. Repository layout

```
app/parsers/blocks.py          Block dataclass and to_dict()
app/parsers/docx_parser.py     the parser, heading/list/table logic, CLI
app/parsers/registry.py        ParserRegistry and build_default_registry()
app/parsers/__init__.py        public exports
tools/make_samples.py          regenerates every sample .docx below
tools/make_snapshots.py        regenerates samples/expected/*.json
tools/parse_docx.py            CLI: print the blocks of one .docx as JSON
samples/demo-a.docx            Chinese 标题 1 / 标题 2 sample
samples/demo-b.docx            English Heading 1 / Heading 2 sample
samples/expected/demo-a.json   expected blocks for demo-a
samples/expected/demo-b.json   expected blocks for demo-b
samples/edge/empty.docx                      empty paragraphs only
samples/edge/header-footer-comment-only.docx header, footer, comment only
samples/edge/image-only.docx                 one picture, no text
tests/test_docx_parser.py      28 test cases
tests/conftest.py              fixtures
tests/reports/                 pytest logs quoted above
```

Regenerate the samples and snapshots after an intentional change:

```bash
python tools/make_samples.py
python tools/make_snapshots.py
python -m pytest -q
```

---

## 6. Differences from the bid proposal

The proposal listed `samples/edge` edge-case documents only implicitly; the
three edge documents above were added so that AC3 has a real file to point at,
since an empty document alone cannot prove that header, footer and comment
text are excluded.

Beyond the proposal, this repository also ships `tools/make_snapshots.py`
(the proposal named only `tools/make_samples.py`) so that the expected-output
snapshots are reproducible in one command as well.

Not in scope for this order: PDF parsing, OCR, image-text extraction, page
numbers, and any persistence of the parsed blocks.
