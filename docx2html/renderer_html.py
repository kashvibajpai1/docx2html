"""
renderer_html.py
----------------
Render a :class:`~docx2html.schema.DocDocument` AST into semantic HTML.

Design goals
~~~~~~~~~~~~
- Zero third-party dependencies beyond the standard library.
- Semantic HTML5 elements (``<h1>``–``<h6>``, ``<ul>``, ``<ol>``, ``<table>``,
  ``<figure>``, ``<code>``, ``<pre>``, ``<hr>``).
- All styling via Tailwind CSS utility classes (no inline ``style=`` attributes).
- Output is readable by humans, diffable by tools, and easy for LLMs to edit.
- The renderer is a pure function of (AST, TailwindMapper) — no side effects.

Entry point::

    from docx2html.renderer_html import render
    html = render(doc)
"""

from __future__ import annotations

import html as _html_lib
from typing import Callable

from docx2html.schema import (
    Block,
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
    TextRun,
)
from docx2html.tailwind_mapper import DEFAULT_MAPPER, TailwindMapper


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def render(
    doc: DocDocument,
    *,
    mapper: TailwindMapper = DEFAULT_MAPPER,
    pretty: bool = True,
    include_document_wrapper: bool = True,
    title: str | None = None,
) -> str:
    """
    Render *doc* to an HTML string.

    Parameters
    ----------
    doc:
        The parsed :class:`~docx2html.schema.DocDocument`.
    mapper:
        Tailwind class mapper.  Pass a custom :class:`~tailwind_mapper.TailwindMapper`
        to override styles.
    pretty:
        When ``True`` (default) indents nested HTML for readability.
        Set to ``False`` for compact single-line output.
    include_document_wrapper:
        When ``True`` (default) wraps the body in a full ``<!DOCTYPE html>``
        page with ``<head>`` containing the Tailwind CDN script.
        Set to ``False`` to emit only the inner body fragment (useful for
        embedding in an existing page).
    title:
        Optional ``<title>`` for the HTML page.  Defaults to
        ``doc.title`` if available, otherwise ``"Document"``.

    Returns
    -------
    str
        The rendered HTML.
    """
    renderer = _HtmlRenderer(mapper=mapper, pretty=pretty)
    body_html = renderer.render_body(doc)

    if not include_document_wrapper:
        return body_html

    page_title = title or doc.title or "Document"
    return _wrap_page(body_html, page_title=page_title, doc_classes=mapper.document)


# ---------------------------------------------------------------------------
# Page wrapper
# ---------------------------------------------------------------------------


def _wrap_page(body_html: str, *, page_title: str, doc_classes: str) -> str:
    """Wrap *body_html* in a complete HTML5 page with Tailwind CDN."""
    return (
        "<!DOCTYPE html>\n"
        '<html lang="en">\n'
        "<head>\n"
        '  <meta charset="UTF-8" />\n'
        '  <meta name="viewport" content="width=device-width, initial-scale=1.0" />\n'
        f"  <title>{_esc(page_title)}</title>\n"
        '  <script src="https://cdn.tailwindcss.com"></script>\n'
        "</head>\n"
        "<body>\n"
        f'  <article class="{doc_classes}">\n'
        f"{_indent(body_html, 4)}\n"
        "  </article>\n"
        "</body>\n"
        "</html>"
    )


# ---------------------------------------------------------------------------
# Renderer class
# ---------------------------------------------------------------------------


class _HtmlRenderer:
    """Stateless renderer that converts AST nodes to HTML strings."""

    def __init__(self, mapper: TailwindMapper, pretty: bool) -> None:
        self._m = mapper
        self._pretty = pretty
        # Dispatch table: AST node type → render method.
        self._dispatch: dict[type, Callable[[object], str]] = {
            Heading: self._render_heading,         # type: ignore[dict-item]
            Paragraph: self._render_paragraph,      # type: ignore[dict-item]
            DocList: self._render_list,             # type: ignore[dict-item]
            Table: self._render_table,              # type: ignore[dict-item]
            Image: self._render_image,              # type: ignore[dict-item]
            CodeBlock: self._render_code_block,     # type: ignore[dict-item]
            HorizontalRule: self._render_hr,        # type: ignore[dict-item]
        }

    # ------------------------------------------------------------------
    # Body
    # ------------------------------------------------------------------

    def render_body(self, doc: DocDocument) -> str:
        parts: list[str] = []
        for block in doc.blocks:
            rendered = self._render_block(block)
            if rendered:
                parts.append(rendered)
        sep = "\n" if self._pretty else ""
        return sep.join(parts)

    def _render_block(self, block: Block) -> str:
        handler = self._dispatch.get(type(block))
        if handler is None:
            # Unknown block type — emit an HTML comment so it's not silently lost.
            return f"<!-- unsupported block: {type(block).__name__} -->"
        return handler(block)  # type: ignore[arg-type]

    # ------------------------------------------------------------------
    # Heading
    # ------------------------------------------------------------------

    def _render_heading(self, node: Heading) -> str:
        tag = f"h{node.level}"
        classes = self._m.heading(node.level)
        inner = self._render_runs(node.runs)
        return f'<{tag} class="{classes}">{inner}</{tag}>'

    # ------------------------------------------------------------------
    # Paragraph
    # ------------------------------------------------------------------

    def _render_paragraph(self, node: Paragraph) -> str:
        if node.is_empty():
            return f'<p class="{self._m.paragraph_empty}">&nbsp;</p>'

        classes = self._m.paragraph
        align_class = self._m.alignment(node.alignment)
        if align_class:
            classes = self._m.merge(classes, align_class)

        inner = self._render_runs(node.runs)
        return f'<p class="{classes}">{inner}</p>'

    # ------------------------------------------------------------------
    # Lists
    # ------------------------------------------------------------------

    def _render_list(self, node: DocList) -> str:
        """
        Render a :class:`~schema.DocList` as nested ``<ul>`` / ``<ol>`` elements.

        Items with ``depth > 0`` are grouped into sub-lists under the preceding
        depth-0 (or shallower) item.
        """
        tag = "ol" if node.ordered else "ul"
        base_classes = self._m.list_ordered if node.ordered else self._m.list_unordered
        return self._render_list_items(node.items, tag=tag, base_classes=base_classes, depth=0)

    def _render_list_items(
        self,
        items: list[ListItem],
        *,
        tag: str,
        base_classes: str,
        depth: int,
    ) -> str:
        """Recursively render list items, grouping nested depths into sub-lists."""
        li_parts: list[str] = []
        i = 0
        while i < len(items):
            item = items[i]
            if item.depth > depth:
                # This item belongs to a deeper level — skip, it will be
                # collected as a sub-list under the previous item.
                i += 1
                continue

            # Collect all immediately following items that are deeper.
            sub_items: list[ListItem] = []
            j = i + 1
            while j < len(items) and items[j].depth > depth:
                sub_items.append(items[j])
                j += 1

            inner_html = self._render_runs(item.runs)

            if sub_items:
                # Render the sub-list and append it inside this <li>.
                sub_list_html = self._render_list_items(
                    sub_items, tag=tag, base_classes=base_classes, depth=depth + 1
                )
                li_content = f"{inner_html}\n{sub_list_html}"
            else:
                li_content = inner_html

            li_parts.append(f'  <li class="{self._m.list_item}">{li_content}</li>')
            i = j  # skip past sub-items we already processed

        items_html = "\n".join(li_parts)

        indent_extra = self._m.list_indent(depth)
        classes = self._m.merge(base_classes, indent_extra) if indent_extra else base_classes

        if self._pretty:
            return f'<{tag} class="{classes}">\n{items_html}\n</{tag}>'
        return f'<{tag} class="{classes}">{items_html}</{tag}>'

    # ------------------------------------------------------------------
    # Table
    # ------------------------------------------------------------------

    def _render_table(self, node: Table) -> str:
        rows_html = self._render_table_rows(node)
        table_html = (
            f'<table class="{self._m.table}">\n'
            f"{rows_html}\n"
            f"</table>"
        )
        return f'<div class="{self._m.table_wrapper}">\n{_indent(table_html, 2)}\n</div>'

    def _render_table_rows(self, node: Table) -> str:
        parts: list[str] = []
        for row_idx, row in enumerate(node.rows):
            is_header_row = node.has_header and row_idx == 0
            row_class = self._m.table_head_row if is_header_row else (
                self._m.table_row_even if row_idx % 2 == 0 else self._m.table_row_odd
            )

            cells_html = self._render_table_cells(row.cells, is_header_row=is_header_row)

            if is_header_row:
                parts.append(
                    f"  <thead>\n"
                    f'    <tr class="{row_class}">\n'
                    f"{cells_html}\n"
                    f"    </tr>\n"
                    f"  </thead>"
                )
            else:
                if row_idx == 1:
                    parts.append("  <tbody>")
                parts.append(
                    f'    <tr class="{row_class}">\n'
                    f"{cells_html}\n"
                    f"    </tr>"
                )

        if len(node.rows) > (1 if node.has_header else 0):
            parts.append("  </tbody>")

        return "\n".join(parts)

    def _render_table_cells(
        self, cells: list[TableCell], *, is_header_row: bool
    ) -> str:
        cell_parts: list[str] = []
        for cell in cells:
            tag = "th" if is_header_row or cell.is_header else "td"
            classes = (
                self._m.table_header_cell
                if (is_header_row or cell.is_header)
                else self._m.table_body_cell
            )
            col_span_attr = f' colspan="{cell.col_span}"' if cell.col_span > 1 else ""
            inner = self._render_runs(cell.runs) if cell.runs else "&nbsp;"
            cell_parts.append(
                f'      <{tag} class="{classes}"{col_span_attr}>{inner}</{tag}>'
            )
        return "\n".join(cell_parts)

    # ------------------------------------------------------------------
    # Image
    # ------------------------------------------------------------------

    def _render_image(self, node: Image) -> str:
        alt = _esc(node.alt or "")
        size_attrs = ""
        if node.width:
            size_attrs += f' width="{node.width}"'
        if node.height:
            size_attrs += f' height="{node.height}"'
        img = (
            f'<img src="{_esc(node.src)}" alt="{alt}"'
            f' class="{self._m.image}"{size_attrs} />'
        )
        return f'<figure class="{self._m.image_wrapper}">\n  {img}\n</figure>'

    # ------------------------------------------------------------------
    # Code block
    # ------------------------------------------------------------------

    def _render_code_block(self, node: CodeBlock) -> str:
        lang_attr = f' data-language="{_esc(node.language)}"' if node.language else ""
        code_html = _esc(node.text)
        code_tag = f'<code class="{self._m.code_block}"{lang_attr}>{code_html}</code>'
        return (
            f'<div class="{self._m.code_block_wrapper}">\n'
            f"  <pre>{code_tag}</pre>\n"
            f"</div>"
        )

    # ------------------------------------------------------------------
    # Horizontal rule
    # ------------------------------------------------------------------

    def _render_hr(self, _node: HorizontalRule) -> str:
        return f'<hr class="{self._m.horizontal_rule}" />'

    # ------------------------------------------------------------------
    # Inline runs
    # ------------------------------------------------------------------

    def _render_runs(self, runs: list[TextRun]) -> str:
        """
        Render a list of :class:`~schema.TextRun` objects to an HTML inline string.

        Consecutive runs with identical styles are *not* merged (kept simple
        and deterministic).
        """
        parts: list[str] = []
        for run in runs:
            parts.append(self._render_run(run))
        return "".join(parts)

    def _render_run(self, run: TextRun) -> str:
        """Wrap a single :class:`~schema.TextRun` in appropriate inline tags."""
        text = _esc(run.text)

        if run.is_plain():
            return text

        # Build up from innermost to outermost tag.
        content = text

        if run.code:
            content = f'<code class="{self._m.inline_code}">{content}</code>'
        else:
            # Apply inline formatting via <span> with combined classes.
            classes: list[str] = []
            if run.bold:
                classes.append(self._m.inline_bold)
            if run.italic:
                classes.append(self._m.inline_italic)
            if run.underline:
                classes.append(self._m.inline_underline)
            if run.strikethrough:
                classes.append(self._m.inline_strikethrough)

            if classes:
                class_str = " ".join(classes)
                content = f'<span class="{class_str}">{content}</span>'

        if run.hyperlink:
            url = _esc(run.hyperlink)
            content = (
                f'<a href="{url}" class="{self._m.inline_link}"'
                f' target="_blank" rel="noopener noreferrer">{content}</a>'
            )

        return content


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------


def _esc(text: str) -> str:
    """HTML-escape *text* to prevent XSS / malformed markup."""
    return _html_lib.escape(text, quote=True)


def _indent(text: str, spaces: int) -> str:
    """Indent every non-empty line in *text* by *spaces* spaces."""
    pad = " " * spaces
    return "\n".join(
        pad + line if line.strip() else line for line in text.splitlines()
    )
