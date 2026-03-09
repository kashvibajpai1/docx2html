"""
test_renderer.py
----------------
Unit tests for docx2html.renderer_html.

These tests construct Document AST nodes directly (no file I/O) and verify
that the HTML output contains the expected elements and Tailwind classes.
"""

from __future__ import annotations

import pytest

from docx2html.renderer_html import render
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
)
from docx2html.tailwind_mapper import TailwindMapper


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _doc(*blocks):
    """Convenience factory: build a DocDocument from positional block args."""
    return DocDocument(blocks=list(blocks))


def _render_fragment(*blocks, mapper=None, pretty=True) -> str:
    """Render blocks as an HTML fragment (no page wrapper)."""
    doc = _doc(*blocks)
    kwargs = dict(include_document_wrapper=False, pretty=pretty)
    if mapper:
        kwargs["mapper"] = mapper
    return render(doc, **kwargs)


# ---------------------------------------------------------------------------
# Full page wrapper
# ---------------------------------------------------------------------------


class TestPageWrapper:
    def test_doctype_present(self):
        html = render(_doc(), pretty=False)
        assert html.startswith("<!DOCTYPE html>")

    def test_tailwind_cdn_present(self):
        html = render(_doc())
        assert "cdn.tailwindcss.com" in html

    def test_custom_title(self):
        html = render(_doc(), title="My Report")
        assert "<title>My Report</title>" in html

    def test_fragment_mode_no_doctype(self):
        html = _render_fragment(Paragraph(runs=[TextRun("hi")]))
        assert "<!DOCTYPE html>" not in html
        assert "<p" in html


# ---------------------------------------------------------------------------
# Headings
# ---------------------------------------------------------------------------


class TestHeadingRendering:
    @pytest.mark.parametrize("level", [1, 2, 3, 4, 5, 6])
    def test_heading_tag(self, level):
        html = _render_fragment(Heading(level=level, runs=[TextRun("Title")]))
        assert f"<h{level} " in html
        assert f"</h{level}>" in html

    def test_heading_content(self):
        html = _render_fragment(Heading(level=1, runs=[TextRun("Hello World")]))
        assert "Hello World" in html

    def test_heading_has_class(self):
        html = _render_fragment(Heading(level=1, runs=[TextRun("T")]))
        assert 'class="' in html
        assert "font-bold" in html

    def test_heading_xss_escaping(self):
        html = _render_fragment(Heading(level=1, runs=[TextRun("<script>alert(1)</script>")]))
        assert "<script>" not in html
        assert "&lt;script&gt;" in html


# ---------------------------------------------------------------------------
# Paragraphs
# ---------------------------------------------------------------------------


class TestParagraphRendering:
    def test_plain_paragraph(self):
        html = _render_fragment(Paragraph(runs=[TextRun("Hello")]))
        assert "<p " in html
        assert "Hello" in html

    def test_empty_paragraph_renders_nbsp(self):
        html = _render_fragment(Paragraph(runs=[TextRun("   ")]))
        assert "&nbsp;" in html

    def test_bold_run(self):
        html = _render_fragment(
            Paragraph(runs=[TextRun("normal "), TextRun("bold", bold=True)])
        )
        assert "font-bold" in html
        assert "<span" in html

    def test_italic_run(self):
        html = _render_fragment(
            Paragraph(runs=[TextRun("text", italic=True)])
        )
        assert "italic" in html

    def test_underline_run(self):
        html = _render_fragment(Paragraph(runs=[TextRun("u", underline=True)]))
        assert "underline" in html

    def test_strikethrough_run(self):
        html = _render_fragment(Paragraph(runs=[TextRun("s", strikethrough=True)]))
        assert "line-through" in html

    def test_inline_code_run(self):
        html = _render_fragment(Paragraph(runs=[TextRun("code", code=True)]))
        assert "<code" in html
        assert "font-mono" in html

    def test_hyperlink_run(self):
        html = _render_fragment(
            Paragraph(runs=[TextRun("click", hyperlink="https://example.com")])
        )
        assert '<a href="https://example.com"' in html
        assert "target=\"_blank\"" in html

    def test_alignment_class_added(self):
        html = _render_fragment(Paragraph(runs=[TextRun("center")], alignment="center"))
        assert "text-center" in html

    def test_special_chars_escaped(self):
        html = _render_fragment(Paragraph(runs=[TextRun('a & b "c"')]))
        assert "&amp;" in html
        assert "&quot;" in html


# ---------------------------------------------------------------------------
# Lists
# ---------------------------------------------------------------------------


class TestListRendering:
    def test_unordered_list(self):
        lst = DocList(
            ordered=False,
            items=[ListItem(runs=[TextRun("A")]), ListItem(runs=[TextRun("B")])],
        )
        html = _render_fragment(lst)
        assert "<ul " in html
        assert "<li " in html
        assert "list-disc" in html

    def test_ordered_list(self):
        lst = DocList(
            ordered=True,
            items=[ListItem(runs=[TextRun("1")]), ListItem(runs=[TextRun("2")])],
        )
        html = _render_fragment(lst)
        assert "<ol " in html
        assert "list-decimal" in html

    def test_list_items_content(self):
        lst = DocList(
            ordered=False,
            items=[ListItem(runs=[TextRun("Alpha")]), ListItem(runs=[TextRun("Beta")])],
        )
        html = _render_fragment(lst)
        assert "Alpha" in html
        assert "Beta" in html

    def test_nested_list(self):
        lst = DocList(
            ordered=False,
            items=[
                ListItem(runs=[TextRun("Parent")], depth=0),
                ListItem(runs=[TextRun("Child")], depth=1),
            ],
        )
        html = _render_fragment(lst)
        # The child should produce a sub-list
        assert html.count("<ul") >= 2


# ---------------------------------------------------------------------------
# Tables
# ---------------------------------------------------------------------------


class TestTableRendering:
    def _make_table(self) -> Table:
        return Table(
            rows=[
                TableRow(
                    cells=[
                        TableCell(runs=[TextRun("Name")], is_header=True),
                        TableCell(runs=[TextRun("Age")], is_header=True),
                    ]
                ),
                TableRow(
                    cells=[
                        TableCell(runs=[TextRun("Alice")]),
                        TableCell(runs=[TextRun("30")]),
                    ]
                ),
            ],
            has_header=True,
        )

    def test_table_element_present(self):
        html = _render_fragment(self._make_table())
        assert "<table " in html

    def test_thead_and_tbody(self):
        html = _render_fragment(self._make_table())
        assert "<thead>" in html
        assert "<tbody>" in html

    def test_th_for_header_cells(self):
        html = _render_fragment(self._make_table())
        assert "<th " in html

    def test_td_for_body_cells(self):
        html = _render_fragment(self._make_table())
        assert "<td " in html

    def test_cell_content(self):
        html = _render_fragment(self._make_table())
        assert "Alice" in html
        assert "30" in html

    def test_table_wrapper_div(self):
        html = _render_fragment(self._make_table())
        assert "overflow-x-auto" in html


# ---------------------------------------------------------------------------
# Images
# ---------------------------------------------------------------------------


class TestImageRendering:
    def test_img_tag(self):
        html = _render_fragment(Image(src="media/img.png", alt="Test image"))
        assert "<img " in html
        assert 'src="media/img.png"' in html
        assert 'alt="Test image"' in html

    def test_figure_wrapper(self):
        html = _render_fragment(Image(src="media/img.png"))
        assert "<figure " in html

    def test_image_class(self):
        html = _render_fragment(Image(src="media/img.png"))
        assert "max-w-full" in html

    def test_width_height_attrs(self):
        html = _render_fragment(Image(src="media/img.png", width=640, height=480))
        assert 'width="640"' in html
        assert 'height="480"' in html


# ---------------------------------------------------------------------------
# Code blocks
# ---------------------------------------------------------------------------


class TestCodeBlockRendering:
    def test_pre_code_tags(self):
        html = _render_fragment(CodeBlock(text='print("hello")', language="python"))
        assert "<pre>" in html
        assert "<code " in html

    def test_language_data_attr(self):
        html = _render_fragment(CodeBlock(text="x=1", language="python"))
        assert 'data-language="python"' in html

    def test_code_content_escaped(self):
        html = _render_fragment(CodeBlock(text="<html>&amp;</html>"))
        assert "&lt;html&gt;" in html

    def test_monospace_class(self):
        html = _render_fragment(CodeBlock(text="x"))
        assert "font-mono" in html


# ---------------------------------------------------------------------------
# Horizontal rule
# ---------------------------------------------------------------------------


class TestHorizontalRuleRendering:
    def test_hr_tag(self):
        html = _render_fragment(HorizontalRule())
        assert "<hr " in html

    def test_hr_class(self):
        html = _render_fragment(HorizontalRule())
        assert "border-t" in html


# ---------------------------------------------------------------------------
# Custom mapper
# ---------------------------------------------------------------------------


class TestCustomMapper:
    def test_custom_heading_class(self):
        from docx2html.tailwind_mapper import HeadingClasses

        mapper = TailwindMapper(
            headings=HeadingClasses(
                h1="custom-h1",
                h2="custom-h2",
                h3="custom-h3",
                h4="custom-h4",
                h5="custom-h5",
                h6="custom-h6",
            )
        )
        html = _render_fragment(Heading(level=1, runs=[TextRun("T")]), mapper=mapper)
        assert "custom-h1" in html

    def test_custom_paragraph_class(self):
        mapper = TailwindMapper(paragraph="my-para-class")
        html = _render_fragment(Paragraph(runs=[TextRun("text")]), mapper=mapper)
        assert "my-para-class" in html


# ---------------------------------------------------------------------------
# Mixed document
# ---------------------------------------------------------------------------


class TestMixedDocument:
    def test_multiple_blocks_all_rendered(self):
        doc = _doc(
            Heading(level=1, runs=[TextRun("Title")]),
            Paragraph(runs=[TextRun("Intro.")]),
            DocList(ordered=False, items=[ListItem(runs=[TextRun("Item")])]),
            HorizontalRule(),
        )
        html = render(doc, include_document_wrapper=False)
        assert "<h1" in html
        assert "<p " in html
        assert "<ul" in html
        assert "<hr" in html
