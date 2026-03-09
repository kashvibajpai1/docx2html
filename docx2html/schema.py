"""
schema.py
---------
Document AST (Abstract Syntax Tree) definitions.

Every element parsed from a .docx file is represented as a typed dataclass
that forms a tree rooted at :class:`DocDocument`.  The AST is the single
source of truth that both the HTML renderer and the JSON exporter consume.

Design goals
~~~~~~~~~~~~
- Pure data, no logic.
- Serialisable to / from plain dicts so JSON export is trivial.
- Extensible: adding a new node type only requires a new dataclass + updating
  the ``Block`` union and ``block_from_dict`` factory.
"""

from __future__ import annotations

import dataclasses
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Union


# ---------------------------------------------------------------------------
# Inline content
# ---------------------------------------------------------------------------


class FontStyle(str, Enum):
    """Inline typographic emphasis flags."""

    NORMAL = "normal"
    BOLD = "bold"
    ITALIC = "italic"
    BOLD_ITALIC = "bold_italic"
    UNDERLINE = "underline"
    STRIKETHROUGH = "strikethrough"
    CODE = "code"  # monospace / inline-code run


@dataclass
class TextRun:
    """
    A contiguous run of text that shares the same inline style.

    A single paragraph is composed of one or more :class:`TextRun` objects,
    which lets us preserve bold, italic, underline, and hyperlink spans.
    """

    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strikethrough: bool = False
    code: bool = False  # monospace run
    hyperlink: str | None = None  # URL if this run is a hyperlink anchor

    def is_plain(self) -> bool:
        """Return True when no inline formatting is applied."""
        return not any(
            [
                self.bold,
                self.italic,
                self.underline,
                self.strikethrough,
                self.code,
                self.hyperlink,
            ]
        )

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {"text": self.text}
        if self.bold:
            d["bold"] = True
        if self.italic:
            d["italic"] = True
        if self.underline:
            d["underline"] = True
        if self.strikethrough:
            d["strikethrough"] = True
        if self.code:
            d["code"] = True
        if self.hyperlink:
            d["hyperlink"] = self.hyperlink
        return d

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "TextRun":
        return cls(
            text=data["text"],
            bold=data.get("bold", False),
            italic=data.get("italic", False),
            underline=data.get("underline", False),
            strikethrough=data.get("strikethrough", False),
            code=data.get("code", False),
            hyperlink=data.get("hyperlink"),
        )


# ---------------------------------------------------------------------------
# Block-level nodes
# ---------------------------------------------------------------------------


@dataclass
class Heading:
    """A section heading with a hierarchical level (1–6)."""

    level: int  # 1 = h1 … 6 = h6
    runs: list[TextRun] = field(default_factory=list)

    @property
    def text(self) -> str:
        """Plain-text content of the heading (no markup)."""
        return "".join(r.text for r in self.runs)

    def to_dict(self) -> dict[str, Any]:
        return {
            "type": "heading",
            "level": self.level,
            "text": self.text,
            "runs": [r.to_dict() for r in self.runs],
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "Heading":
        runs = [TextRun.from_dict(r) for r in data.get("runs", [])]
        if not runs and "text" in data:
            runs = [TextRun(text=data["text"])]
        return cls(level=data["level"], runs=runs)


@dataclass
class Paragraph:
    """A body paragraph, potentially containing rich inline runs."""

    runs: list[TextRun] = field(default_factory=list)
    # Alignment hint preserved from the source document.
    alignment: str | None = None  # "left" | "center" | "right" | "justify"

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)

    def is_empty(self) -> bool:
        return not self.text.strip()

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "type": "paragraph",
            "text": self.text,
            "runs": [r.to_dict() for r in self.runs],
        }
        if self.alignment:
            d["alignment"] = self.alignment
        return d

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "Paragraph":
        runs = [TextRun.from_dict(r) for r in data.get("runs", [])]
        if not runs and "text" in data:
            runs = [TextRun(text=data["text"])]
        return cls(runs=runs, alignment=data.get("alignment"))


@dataclass
class ListItem:
    """A single item inside a :class:`DocList`."""

    runs: list[TextRun] = field(default_factory=list)
    # Nesting depth (0 = top-level).
    depth: int = 0

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)

    def to_dict(self) -> dict[str, Any]:
        return {
            "text": self.text,
            "runs": [r.to_dict() for r in self.runs],
            "depth": self.depth,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "ListItem":
        runs = [TextRun.from_dict(r) for r in data.get("runs", [])]
        if not runs and "text" in data:
            runs = [TextRun(text=data["text"])]
        return cls(runs=runs, depth=data.get("depth", 0))


@dataclass
class DocList:
    """
    An ordered or unordered list.

    Items may have a ``depth`` field for nested sub-lists; the renderer is
    responsible for grouping them into nested ``<ul>`` / ``<ol>`` elements.
    """

    ordered: bool
    items: list[ListItem] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "type": "list",
            "ordered": self.ordered,
            "items": [i.to_dict() for i in self.items],
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "DocList":
        return cls(
            ordered=data["ordered"],
            items=[ListItem.from_dict(i) for i in data.get("items", [])],
        )


@dataclass
class TableCell:
    """A single cell inside a :class:`TableRow`."""

    runs: list[TextRun] = field(default_factory=list)
    # How many columns this cell spans (from the DOCX grid).
    col_span: int = 1
    # True when this cell is part of the header row.
    is_header: bool = False

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "text": self.text,
            "runs": [r.to_dict() for r in self.runs],
        }
        if self.col_span != 1:
            d["col_span"] = self.col_span
        if self.is_header:
            d["is_header"] = True
        return d

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "TableCell":
        runs = [TextRun.from_dict(r) for r in data.get("runs", [])]
        if not runs and "text" in data:
            runs = [TextRun(text=data["text"])]
        return cls(
            runs=runs,
            col_span=data.get("col_span", 1),
            is_header=data.get("is_header", False),
        )


@dataclass
class TableRow:
    """A row of :class:`TableCell` objects."""

    cells: list[TableCell] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {"cells": [c.to_dict() for c in self.cells]}

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "TableRow":
        return cls(cells=[TableCell.from_dict(c) for c in data.get("cells", [])])


@dataclass
class Table:
    """
    A rectangular grid of cells.

    The first row is treated as a header row when ``has_header`` is True.
    """

    rows: list[TableRow] = field(default_factory=list)
    has_header: bool = True

    def to_dict(self) -> dict[str, Any]:
        return {
            "type": "table",
            "has_header": self.has_header,
            "rows": [r.to_dict() for r in self.rows],
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "Table":
        return cls(
            rows=[TableRow.from_dict(r) for r in data.get("rows", [])],
            has_header=data.get("has_header", True),
        )


@dataclass
class Image:
    """An embedded image extracted from the document."""

    src: str  # Relative path to the saved image file, e.g. "media/image1.png"
    alt: str = ""
    width: int | None = None   # pixels, if available
    height: int | None = None  # pixels, if available

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {"type": "image", "src": self.src, "alt": self.alt}
        if self.width is not None:
            d["width"] = self.width
        if self.height is not None:
            d["height"] = self.height
        return d

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "Image":
        return cls(
            src=data["src"],
            alt=data.get("alt", ""),
            width=data.get("width"),
            height=data.get("height"),
        )


@dataclass
class CodeBlock:
    """A fenced code block (e.g. from a ``Code`` or ``Verbatim`` paragraph style)."""

    text: str
    language: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {"type": "code_block", "language": self.language, "text": self.text}

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "CodeBlock":
        return cls(text=data["text"], language=data.get("language", ""))


@dataclass
class HorizontalRule:
    """A thematic break / horizontal rule (``<hr>``)."""

    def to_dict(self) -> dict[str, Any]:
        return {"type": "horizontal_rule"}

    @classmethod
    def from_dict(cls, _data: dict[str, Any]) -> "HorizontalRule":
        return cls()


# ---------------------------------------------------------------------------
# Union type & factory
# ---------------------------------------------------------------------------

#: All supported block-level node types.
Block = Union[
    Heading,
    Paragraph,
    DocList,
    Table,
    Image,
    CodeBlock,
    HorizontalRule,
]

_BLOCK_TYPE_MAP: dict[str, type] = {
    "heading": Heading,
    "paragraph": Paragraph,
    "list": DocList,
    "table": Table,
    "image": Image,
    "code_block": CodeBlock,
    "horizontal_rule": HorizontalRule,
}


def block_from_dict(data: dict[str, Any]) -> Block:
    """Deserialise a block node from its dict representation."""
    node_type = data.get("type", "")
    klass = _BLOCK_TYPE_MAP.get(node_type)
    if klass is None:
        raise ValueError(f"Unknown block type: {node_type!r}")
    return klass.from_dict(data)  # type: ignore[attr-defined]


# ---------------------------------------------------------------------------
# Root document node
# ---------------------------------------------------------------------------


@dataclass
class DocDocument:
    """
    The root node of the Document AST.

    Contains an ordered sequence of :class:`Block` nodes that represent the
    full content of the source ``.docx`` file.
    """

    blocks: list[Block] = field(default_factory=list)
    # Optional metadata extracted from document core properties.
    title: str = ""
    author: str = ""
    created: str = ""  # ISO-8601 string

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "type": "document",
            "blocks": [b.to_dict() for b in self.blocks],
        }
        if self.title:
            d["title"] = self.title
        if self.author:
            d["author"] = self.author
        if self.created:
            d["created"] = self.created
        return d

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "DocDocument":
        return cls(
            blocks=[block_from_dict(b) for b in data.get("blocks", [])],
            title=data.get("title", ""),
            author=data.get("author", ""),
            created=data.get("created", ""),
        )
