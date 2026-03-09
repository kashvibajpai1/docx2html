"""
tailwind_mapper.py
------------------
Map Document AST node types to Tailwind CSS utility class strings.

Design notes
~~~~~~~~~~~~
- All mappings are pure data — no logic, no conditionals beyond simple lookups.
- The :class:`TailwindMapper` class is instantiated once and shared across the
  renderer.  Callers can subclass or replace it entirely to customise styling
  without touching the renderer.
- Class strings follow the same ordering convention as the official Tailwind
  docs: layout → typography → spacing → borders → effects.
- Dark-mode variants are excluded intentionally to keep the output simple and
  LLM-editable.  A future ``dark_mode=True`` constructor flag could add them.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import NamedTuple


class HeadingClasses(NamedTuple):
    """Tailwind classes for each heading level h1–h6."""

    h1: str
    h2: str
    h3: str
    h4: str
    h5: str
    h6: str

    def get(self, level: int) -> str:
        """Return classes for heading *level* (1-indexed, clamped to 1–6)."""
        level = max(1, min(6, level))
        return getattr(self, f"h{level}")


@dataclass
class TailwindMapper:
    """
    A configurable mapping from AST node types to Tailwind CSS class strings.

    All attributes are plain strings (space-separated Tailwind utility classes)
    so they can be overridden at construction time or patched after the fact.

    Example — override paragraph spacing::

        mapper = TailwindMapper(paragraph="text-base mb-4 leading-relaxed")
    """

    # ------------------------------------------------------------------
    # Headings
    # ------------------------------------------------------------------
    headings: HeadingClasses = field(
        default_factory=lambda: HeadingClasses(
            h1="text-3xl font-bold mb-4 mt-6 leading-tight text-gray-900",
            h2="text-2xl font-semibold mb-3 mt-5 leading-snug text-gray-800",
            h3="text-xl font-semibold mb-2 mt-4 leading-snug text-gray-800",
            h4="text-lg font-medium mb-2 mt-3 text-gray-700",
            h5="text-base font-medium mb-1 mt-2 text-gray-700",
            h6="text-sm font-medium mb-1 mt-2 text-gray-600 uppercase tracking-wide",
        )
    )

    # ------------------------------------------------------------------
    # Body text
    # ------------------------------------------------------------------
    paragraph: str = "text-base mb-3 leading-relaxed text-gray-800"

    #: Empty / spacer paragraph (rendered as a small vertical gap).
    paragraph_empty: str = "mb-2"

    # ------------------------------------------------------------------
    # Inline text formatting
    # ------------------------------------------------------------------
    inline_bold: str = "font-bold"
    inline_italic: str = "italic"
    inline_underline: str = "underline"
    inline_strikethrough: str = "line-through"
    inline_code: str = (
        "font-mono text-sm bg-gray-100 text-red-600 px-1 py-0.5 rounded"
    )
    inline_link: str = "text-blue-600 underline hover:text-blue-800"

    # ------------------------------------------------------------------
    # Lists
    # ------------------------------------------------------------------
    list_unordered: str = "list-disc ml-6 mb-3 space-y-1 text-gray-800"
    list_ordered: str = "list-decimal ml-6 mb-3 space-y-1 text-gray-800"
    list_item: str = "text-base leading-relaxed"
    #: Additional left margin applied per nesting level.
    list_indent_step: str = "ml-4"

    # ------------------------------------------------------------------
    # Tables
    # ------------------------------------------------------------------
    table_wrapper: str = "overflow-x-auto mb-4"
    table: str = "table-auto w-full border-collapse border border-gray-300 text-sm"
    table_head_row: str = "bg-gray-100"
    table_header_cell: str = (
        "border border-gray-300 px-4 py-2 font-semibold text-left text-gray-700"
    )
    table_body_cell: str = "border border-gray-300 px-4 py-2 text-gray-800"
    table_row_even: str = "bg-white"
    table_row_odd: str = "bg-gray-50"

    # ------------------------------------------------------------------
    # Images
    # ------------------------------------------------------------------
    image_wrapper: str = "my-4"
    image: str = "max-w-full h-auto rounded-lg shadow-sm"

    # ------------------------------------------------------------------
    # Code blocks
    # ------------------------------------------------------------------
    code_block_wrapper: str = "mb-4"
    code_block: str = (
        "block font-mono text-sm bg-gray-900 text-green-400 "
        "p-4 rounded-lg overflow-x-auto whitespace-pre"
    )

    # ------------------------------------------------------------------
    # Horizontal rule
    # ------------------------------------------------------------------
    horizontal_rule: str = "border-t border-gray-300 my-6"

    # ------------------------------------------------------------------
    # Document wrapper
    # ------------------------------------------------------------------
    document: str = (
        "max-w-4xl mx-auto px-6 py-8 font-sans text-gray-900 bg-white"
    )

    # ------------------------------------------------------------------
    # Alignment helpers
    # ------------------------------------------------------------------
    _ALIGNMENT_CLASSES: dict[str, str] = field(
        default_factory=lambda: {
            "left": "text-left",
            "center": "text-center",
            "right": "text-right",
            "justify": "text-justify",
        },
        repr=False,
    )

    # ------------------------------------------------------------------
    # Public helper methods
    # ------------------------------------------------------------------

    def heading(self, level: int) -> str:
        """Return the Tailwind classes for heading *level*."""
        return self.headings.get(level)

    def alignment(self, align: str | None) -> str:
        """Return the Tailwind text-alignment class for *align*, or empty string."""
        if align is None:
            return ""
        return self._ALIGNMENT_CLASSES.get(align, "")

    def list_indent(self, depth: int) -> str:
        """
        Return extra indentation classes for a nested list at *depth* levels.

        ``depth == 0`` → no extra indentation (classes already in ``list_unordered``).
        ``depth == 1`` → one :attr:`list_indent_step`.
        """
        if depth <= 0:
            return ""
        return " ".join([self.list_indent_step] * depth)

    def merge(self, base: str, extra: str) -> str:
        """Combine *base* and *extra* class strings, stripping duplicates."""
        base_classes = base.split()
        extra_classes = [c for c in extra.split() if c not in base_classes]
        return " ".join(base_classes + extra_classes)


# ---------------------------------------------------------------------------
# Default singleton — importable directly for convenience.
# ---------------------------------------------------------------------------

#: A ready-to-use :class:`TailwindMapper` with default classes.
DEFAULT_MAPPER: TailwindMapper = TailwindMapper()
