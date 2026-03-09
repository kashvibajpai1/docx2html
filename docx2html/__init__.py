"""
docx2html
---------
Convert Microsoft Word (.docx) files to clean, structured HTML
optimised for LLM ingestion and editing.

Public API
~~~~~~~~~~
::

    from docx2html import parse, render

    doc = parse("template.docx")           # → DocDocument AST
    html = render(doc)                     # → HTML string
    json_ast = doc.to_dict()               # → dict (JSON-serialisable)

The package also exposes the schema types for type annotations and for
constructing documents programmatically in tests.
"""

from docx2html.parser import parse
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
    block_from_dict,
)
from docx2html.tailwind_mapper import DEFAULT_MAPPER, TailwindMapper

__version__ = "0.1.0"
__all__ = [
    # Core pipeline
    "parse",
    "render",
    # Schema types
    "DocDocument",
    "Heading",
    "Paragraph",
    "DocList",
    "ListItem",
    "Table",
    "TableRow",
    "TableCell",
    "Image",
    "CodeBlock",
    "HorizontalRule",
    "TextRun",
    "block_from_dict",
    # Styling
    "TailwindMapper",
    "DEFAULT_MAPPER",
    # Metadata
    "__version__",
]
