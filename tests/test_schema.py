"""
test_schema.py
--------------
Unit tests for docx2html.schema — Document AST construction and serialisation.
"""

from __future__ import annotations

import pytest

from docx2html.schema import (
    CodeBlock,
    DocDocument,
    DocList,
    Heading,
    HorizontalRule,
    Image,
    ListItem,
    Paragraph,
    Table,
    TableCell,
    TableRow,
    TextRun,
    block_from_dict,
)


# ---------------------------------------------------------------------------
# TextRun
# ---------------------------------------------------------------------------


class TestTextRun:
    def test_plain_run(self):
        run = TextRun(text="hello")
        assert run.text == "hello"
        assert run.is_plain() is True

    def test_bold_run_not_plain(self):
        run = TextRun(text="bold", bold=True)
        assert run.is_plain() is False

    def test_to_dict_plain(self):
        run = TextRun(text="hello")
        d = run.to_dict()
        assert d == {"text": "hello"}

    def test_to_dict_formatted(self):
        run = TextRun(text="link", bold=True, hyperlink="https://example.com")
        d = run.to_dict()
        assert d["bold"] is True
        assert d["hyperlink"] == "https://example.com"
        # Keys with False values should be omitted
        assert "italic" not in d

    def test_round_trip(self):
        run = TextRun(text="test", italic=True, underline=True)
        assert TextRun.from_dict(run.to_dict()) == run


# ---------------------------------------------------------------------------
# Heading
# ---------------------------------------------------------------------------


class TestHeading:
    def test_text_property(self):
        heading = Heading(level=1, runs=[TextRun("Title")])
        assert heading.text == "Title"

    def test_to_dict(self):
        heading = Heading(level=2, runs=[TextRun("Sub")])
        d = heading.to_dict()
        assert d["type"] == "heading"
        assert d["level"] == 2
        assert d["text"] == "Sub"

    def test_round_trip(self):
        heading = Heading(level=3, runs=[TextRun("Test", bold=True)])
        restored = Heading.from_dict(heading.to_dict())
        assert restored.level == 3
        assert restored.runs[0].bold is True


# ---------------------------------------------------------------------------
# Paragraph
# ---------------------------------------------------------------------------


class TestParagraph:
    def test_is_empty(self):
        para = Paragraph(runs=[TextRun("   ")])
        assert para.is_empty() is True

    def test_is_not_empty(self):
        para = Paragraph(runs=[TextRun("hello")])
        assert para.is_empty() is False

    def test_to_dict(self):
        para = Paragraph(runs=[TextRun("hello")], alignment="center")
        d = para.to_dict()
        assert d["type"] == "paragraph"
        assert d["alignment"] == "center"

    def test_round_trip(self):
        para = Paragraph(runs=[TextRun("world")], alignment="right")
        restored = Paragraph.from_dict(para.to_dict())
        assert restored.alignment == "right"
        assert restored.text == "world"


# ---------------------------------------------------------------------------
# DocList
# ---------------------------------------------------------------------------


class TestDocList:
    def test_ordered(self):
        lst = DocList(ordered=True, items=[ListItem(runs=[TextRun("one")])])
        d = lst.to_dict()
        assert d["type"] == "list"
        assert d["ordered"] is True

    def test_unordered_round_trip(self):
        lst = DocList(
            ordered=False,
            items=[
                ListItem(runs=[TextRun("A")], depth=0),
                ListItem(runs=[TextRun("B")], depth=1),
            ],
        )
        restored = DocList.from_dict(lst.to_dict())
        assert restored.ordered is False
        assert len(restored.items) == 2
        assert restored.items[1].depth == 1


# ---------------------------------------------------------------------------
# Table
# ---------------------------------------------------------------------------


class TestTable:
    def test_table_round_trip(self):
        table = Table(
            rows=[
                TableRow(cells=[TableCell(runs=[TextRun("Name")], is_header=True)]),
                TableRow(cells=[TableCell(runs=[TextRun("Alice")])]),
            ],
            has_header=True,
        )
        d = table.to_dict()
        assert d["type"] == "table"
        restored = Table.from_dict(d)
        assert len(restored.rows) == 2
        assert restored.rows[0].cells[0].is_header is True


# ---------------------------------------------------------------------------
# Image, CodeBlock, HorizontalRule
# ---------------------------------------------------------------------------


class TestLeafBlocks:
    def test_image_round_trip(self):
        img = Image(src="media/img1.png", alt="Alt text", width=800, height=600)
        restored = Image.from_dict(img.to_dict())
        assert restored.src == "media/img1.png"
        assert restored.width == 800

    def test_code_block_round_trip(self):
        cb = CodeBlock(text='print("hello")', language="python")
        restored = CodeBlock.from_dict(cb.to_dict())
        assert restored.language == "python"
        assert 'print("hello")' in restored.text

    def test_horizontal_rule_round_trip(self):
        hr = HorizontalRule()
        d = hr.to_dict()
        assert d["type"] == "horizontal_rule"
        HorizontalRule.from_dict(d)  # should not raise


# ---------------------------------------------------------------------------
# DocDocument
# ---------------------------------------------------------------------------


class TestDocDocument:
    def test_empty_document(self):
        doc = DocDocument()
        assert doc.blocks == []
        d = doc.to_dict()
        assert d["type"] == "document"
        assert d["blocks"] == []

    def test_round_trip(self):
        doc = DocDocument(
            blocks=[
                Heading(level=1, runs=[TextRun("Hello")]),
                Paragraph(runs=[TextRun("World")]),
            ],
            title="Test Doc",
            author="Alice",
        )
        d = doc.to_dict()
        restored = DocDocument.from_dict(d)
        assert len(restored.blocks) == 2
        assert restored.title == "Test Doc"
        assert isinstance(restored.blocks[0], Heading)
        assert isinstance(restored.blocks[1], Paragraph)


# ---------------------------------------------------------------------------
# block_from_dict factory
# ---------------------------------------------------------------------------


class TestBlockFromDict:
    @pytest.mark.parametrize(
        "block_dict,expected_type",
        [
            ({"type": "heading", "level": 1, "runs": [{"text": "T"}]}, Heading),
            ({"type": "paragraph", "runs": [{"text": "P"}]}, Paragraph),
            ({"type": "list", "ordered": False, "items": []}, DocList),
            ({"type": "table", "has_header": True, "rows": []}, Table),
            ({"type": "image", "src": "a.png", "alt": ""}, Image),
            ({"type": "code_block", "text": "x=1", "language": ""}, CodeBlock),
            ({"type": "horizontal_rule"}, HorizontalRule),
        ],
    )
    def test_dispatch(self, block_dict, expected_type):
        block = block_from_dict(block_dict)
        assert isinstance(block, expected_type)

    def test_unknown_type_raises(self):
        with pytest.raises(ValueError, match="Unknown block type"):
            block_from_dict({"type": "unknown_node"})
