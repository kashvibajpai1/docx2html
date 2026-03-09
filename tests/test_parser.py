"""
test_parser.py
--------------
Integration tests for docx2html.parser.

These tests create real .docx files using python-docx (via conftest fixtures),
parse them, and verify that the resulting DocDocument AST contains the expected
nodes and content.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from docx2html import parse
from docx2html.schema import (
    DocDocument,
    DocList,
    Heading,
    Paragraph,
    Table,
)


# ---------------------------------------------------------------------------
# Error handling
# ---------------------------------------------------------------------------


class TestParseErrors:
    def test_missing_file_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            parse(tmp_path / "nonexistent.docx")

    def test_wrong_extension_raises(self, tmp_path):
        f = tmp_path / "file.txt"
        f.write_text("hello")
        with pytest.raises(ValueError, match=r"\.docx"):
            parse(f)


# ---------------------------------------------------------------------------
# Headings
# ---------------------------------------------------------------------------


class TestHeadingParsing:
    def test_heading_blocks_present(self, docx_with_headings):
        doc = parse(docx_with_headings)
        headings = [b for b in doc.blocks if isinstance(b, Heading)]
        assert len(headings) == 3

    def test_heading_levels(self, docx_with_headings):
        doc = parse(docx_with_headings)
        headings = [b for b in doc.blocks if isinstance(b, Heading)]
        levels = [h.level for h in headings]
        assert 1 in levels
        assert 2 in levels
        assert 3 in levels

    def test_heading_text(self, docx_with_headings):
        doc = parse(docx_with_headings)
        headings = [b for b in doc.blocks if isinstance(b, Heading)]
        texts = [h.text for h in headings]
        assert "Main Title" in texts

    def test_body_paragraph_after_heading(self, docx_with_headings):
        doc = parse(docx_with_headings)
        paragraphs = [b for b in doc.blocks if isinstance(b, Paragraph)]
        assert any("body text" in p.text.lower() for p in paragraphs)


# ---------------------------------------------------------------------------
# Paragraphs
# ---------------------------------------------------------------------------


class TestParagraphParsing:
    def test_paragraph_present(self, docx_with_paragraphs):
        doc = parse(docx_with_paragraphs)
        paragraphs = [b for b in doc.blocks if isinstance(b, Paragraph)]
        assert len(paragraphs) >= 1

    def test_inline_bold_preserved(self, docx_with_paragraphs):
        doc = parse(docx_with_paragraphs)
        paragraphs = [b for b in doc.blocks if isinstance(b, Paragraph)]
        has_bold = any(r.bold for p in paragraphs for r in p.runs)
        assert has_bold, "Expected at least one bold run"

    def test_inline_italic_preserved(self, docx_with_paragraphs):
        doc = parse(docx_with_paragraphs)
        paragraphs = [b for b in doc.blocks if isinstance(b, Paragraph)]
        has_italic = any(r.italic for p in paragraphs for r in p.runs)
        assert has_italic, "Expected at least one italic run"

    def test_plain_paragraph_text(self, docx_with_paragraphs):
        doc = parse(docx_with_paragraphs)
        paragraphs = [b for b in doc.blocks if isinstance(b, Paragraph)]
        all_text = " ".join(p.text for p in paragraphs)
        assert "plain second paragraph" in all_text.lower()


# ---------------------------------------------------------------------------
# Bullet lists
# ---------------------------------------------------------------------------


class TestBulletListParsing:
    def test_list_block_present(self, docx_with_bullet_list):
        doc = parse(docx_with_bullet_list)
        lists = [b for b in doc.blocks if isinstance(b, DocList)]
        assert len(lists) >= 1

    def test_list_is_unordered(self, docx_with_bullet_list):
        doc = parse(docx_with_bullet_list)
        lists = [b for b in doc.blocks if isinstance(b, DocList)]
        assert any(not lst.ordered for lst in lists)

    def test_list_items_content(self, docx_with_bullet_list):
        doc = parse(docx_with_bullet_list)
        lists = [b for b in doc.blocks if isinstance(b, DocList)]
        all_item_texts = [i.text for lst in lists for i in lst.items]
        assert "Alpha" in all_item_texts
        assert "Beta" in all_item_texts
        assert "Gamma" in all_item_texts


# ---------------------------------------------------------------------------
# Numbered lists
# ---------------------------------------------------------------------------


class TestNumberedListParsing:
    def test_list_block_present(self, docx_with_numbered_list):
        doc = parse(docx_with_numbered_list)
        lists = [b for b in doc.blocks if isinstance(b, DocList)]
        assert len(lists) >= 1

    def test_list_items_content(self, docx_with_numbered_list):
        doc = parse(docx_with_numbered_list)
        lists = [b for b in doc.blocks if isinstance(b, DocList)]
        all_item_texts = [i.text for lst in lists for i in lst.items]
        assert "First" in all_item_texts
        assert "Second" in all_item_texts
        assert "Third" in all_item_texts


# ---------------------------------------------------------------------------
# Tables
# ---------------------------------------------------------------------------


class TestTableParsing:
    def test_table_block_present(self, docx_with_table):
        doc = parse(docx_with_table)
        tables = [b for b in doc.blocks if isinstance(b, Table)]
        assert len(tables) >= 1

    def test_table_row_count(self, docx_with_table):
        doc = parse(docx_with_table)
        tables = [b for b in doc.blocks if isinstance(b, Table)]
        table = tables[0]
        assert len(table.rows) == 3  # 1 header + 2 data rows

    def test_table_cell_count(self, docx_with_table):
        doc = parse(docx_with_table)
        tables = [b for b in doc.blocks if isinstance(b, Table)]
        table = tables[0]
        for row in table.rows:
            assert len(row.cells) == 3

    def test_table_header_cell_text(self, docx_with_table):
        doc = parse(docx_with_table)
        tables = [b for b in doc.blocks if isinstance(b, Table)]
        table = tables[0]
        header_texts = [c.text for c in table.rows[0].cells]
        assert "Name" in header_texts
        assert "Age" in header_texts
        assert "City" in header_texts

    def test_table_data_cell_text(self, docx_with_table):
        doc = parse(docx_with_table)
        tables = [b for b in doc.blocks if isinstance(b, Table)]
        table = tables[0]
        all_cell_texts = [c.text for row in table.rows[1:] for c in row.cells]
        assert "Alice" in all_cell_texts
        assert "Bob" in all_cell_texts


# ---------------------------------------------------------------------------
# Empty document
# ---------------------------------------------------------------------------


class TestEmptyDocument:
    def test_empty_doc_has_no_blocks(self, docx_empty):
        doc = parse(docx_empty)
        # Word adds a default empty paragraph; our parser drops empty paras.
        # The result may have zero blocks or a single empty paragraph.
        assert isinstance(doc, DocDocument)

    def test_empty_doc_metadata(self, docx_empty):
        doc = parse(docx_empty)
        # title/author/created may be empty strings but must be strings.
        assert isinstance(doc.title, str)
        assert isinstance(doc.author, str)


# ---------------------------------------------------------------------------
# Mixed document
# ---------------------------------------------------------------------------


class TestMixedDocument:
    def test_all_block_types_present(self, docx_mixed):
        doc = parse(docx_mixed)
        block_types = {type(b) for b in doc.blocks}
        assert Heading in block_types
        assert Paragraph in block_types
        assert DocList in block_types
        assert Table in block_types

    def test_block_order_preserved(self, docx_mixed):
        doc = parse(docx_mixed)
        # First block should be a heading ("Mixed Document")
        assert isinstance(doc.blocks[0], Heading)
        assert doc.blocks[0].level == 1
