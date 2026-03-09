# docx2html

Convert Microsoft Word (`.docx`) files to clean, structured HTML optimised for LLM ingestion and editing.

## The Problem

When you upload a `.docx` template to an LLM and ask it to fill in content, the model receives raw binary or poorly-structured text. Formatting is lost, tables collapse, and the model has no reliable way to preserve the document's structure.

**docx2html** solves this by converting `.docx` files into semantic, Tailwind-styled HTML that LLMs can read, understand, and edit without breaking the document's formatting.

## Workflow

```
template.docx  →  docx2html  →  template.html  →  LLM  →  filled_template.html
```

1. Convert `template.docx` → `template.html`
2. Feed `template.html` to an LLM with your instructions
3. The LLM edits the HTML content while preserving structure and Tailwind classes
4. Render or save the modified HTML as needed

## Architecture

The conversion pipeline has four stages:

```
DOCX
 ↓
python-docx parser       (parser.py)
 ↓
Document AST             (schema.py)
 ↓
HTML Renderer            (renderer_html.py)
 ↓
Tailwind-styled HTML
```

### Document AST (`schema.py`)

Every element in the `.docx` file is mapped to a typed Python dataclass:

| AST Node | DOCX source | HTML output |
|---|---|---|
| `Heading` | `Heading 1`–`Heading 9` styles | `<h1>`–`<h6>` |
| `Paragraph` | Normal text | `<p>` |
| `DocList` | `List Bullet`, `List Number` styles | `<ul>`, `<ol>` |
| `Table` | Word tables | `<table>` with `<thead>`/`<tbody>` |
| `Image` | Embedded images | `<figure><img /></figure>` |
| `CodeBlock` | `Code`, `Verbatim` styles | `<pre><code>` |
| `HorizontalRule` | Paragraph border bottom | `<hr>` |

Inline formatting (`TextRun`) captures bold, italic, underline, strikethrough, monospace, and hyperlinks.

### Parser (`parser.py`)

Uses `python-docx` exclusively — no raw XML manipulation. Iterates the document body in order, grouping consecutive list paragraphs into `DocList` nodes and extracting embedded images to a `media/` directory.

### Renderer (`renderer_html.py`)

Converts the AST to semantic HTML5. No inline `style=` attributes — all styling uses Tailwind CSS utility classes via the mapper.

### Tailwind Mapper (`tailwind_mapper.py`)

A configurable dataclass mapping each AST node type to Tailwind utility class strings. Override any mapping without touching the renderer:

```python
from docx2html.tailwind_mapper import TailwindMapper

mapper = TailwindMapper(
    paragraph="text-lg mb-4 leading-loose text-gray-700",
)
```

## Installation

```bash
pip install docx2html
```

Requires Python 3.10+.

## CLI Usage

```bash
# Convert to HTML (output: template.html in the same directory)
docx2html template.docx

# Explicit output path
docx2html template.docx -o output.html

# Print HTML to stdout
docx2html template.docx --stdout

# Export Document AST as JSON
docx2html template.docx --json

# JSON to stdout
docx2html template.docx --json --stdout

# Compact (non-pretty) HTML
docx2html template.docx --no-pretty

# HTML fragment only (no <!DOCTYPE html> wrapper)
docx2html template.docx --fragment

# Custom page title
docx2html template.docx --title "Q3 Report"

# Custom media directory for extracted images
docx2html template.docx --media-dir ./assets/images

# Verbose debug logging
docx2html template.docx -v

# Show version
docx2html --version
```

## Python API

```python
from docx2html import parse, render

# Parse a .docx file into a Document AST
doc = parse("template.docx")

# Render to HTML
html = render(doc)

# Export AST as JSON-serialisable dict
import json
print(json.dumps(doc.to_dict(), indent=2))

# Custom Tailwind classes
from docx2html.tailwind_mapper import TailwindMapper
mapper = TailwindMapper(paragraph="text-lg mb-6")
html = render(doc, mapper=mapper)

# Fragment mode (no page wrapper)
fragment = render(doc, include_document_wrapper=False)
```

## HTML Output Example

Input `.docx` with a heading, paragraph, and table:

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <title>Project Proposal</title>
  <script src="https://cdn.tailwindcss.com"></script>
</head>
<body>
  <article class="max-w-4xl mx-auto px-6 py-8 font-sans text-gray-900 bg-white">
    <h1 class="text-3xl font-bold mb-4 mt-6 leading-tight text-gray-900">Project Proposal</h1>
    <h2 class="text-2xl font-semibold mb-3 mt-5 leading-snug text-gray-800">Overview</h2>
    <p class="text-base mb-3 leading-relaxed text-gray-800">This document describes the plan.</p>
    <ul class="list-disc ml-6 mb-3 space-y-1 text-gray-800">
      <li class="text-base leading-relaxed">Design spec</li>
      <li class="text-base leading-relaxed">Implementation</li>
    </ul>
    <div class="overflow-x-auto mb-4">
      <table class="table-auto w-full border-collapse border border-gray-300 text-sm">
        <thead>
          <tr class="bg-gray-100">
            <th class="border border-gray-300 px-4 py-2 font-semibold text-left text-gray-700">Phase</th>
            <th class="border border-gray-300 px-4 py-2 font-semibold text-left text-gray-700">Date</th>
          </tr>
        </thead>
        <tbody>
          <tr class="bg-gray-50">
            <td class="border border-gray-300 px-4 py-2 text-gray-800">Alpha</td>
            <td class="border border-gray-300 px-4 py-2 text-gray-800">Jan 1</td>
          </tr>
        </tbody>
      </table>
    </div>
  </article>
</body>
</html>
```

## JSON AST Example

```json
{
  "type": "document",
  "blocks": [
    {"type": "heading", "level": 1, "text": "Project Proposal", "runs": [{"text": "Project Proposal"}]},
    {"type": "paragraph", "text": "This document describes the plan.", "runs": [{"text": "This document describes the plan."}]},
    {
      "type": "list",
      "ordered": false,
      "items": [
        {"text": "Design spec", "runs": [{"text": "Design spec"}], "depth": 0},
        {"text": "Implementation", "runs": [{"text": "Implementation"}], "depth": 0}
      ]
    }
  ]
}
```

## Why This Improves LLM Document Editing

| Problem | Without docx2html | With docx2html |
|---|---|---|
| Structure | Binary DOCX, opaque to LLMs | Semantic HTML the LLM can read |
| Formatting | Lost during extraction | Preserved as Tailwind classes |
| Tables | Flattened or broken | Proper `<thead>`/`<tbody>` |
| Lists | Merged into plain text | `<ul>`/`<ol>` with nesting |
| Images | Lost | Referenced as `<img>` tags |
| Editability | LLM cannot safely modify | LLM edits content, leaves tags intact |

## Project Structure

```
docx2html/
├── __init__.py          # Public API exports
├── schema.py            # Document AST (typed dataclasses)
├── parser.py            # python-docx → AST
├── renderer_html.py     # AST → HTML
├── tailwind_mapper.py   # Element → Tailwind class mapping
├── utils.py             # Shared utilities (I/O, paths, JSON)
└── cli.py               # CLI (typer)

tests/
├── conftest.py          # Shared fixtures (in-memory .docx files)
├── test_schema.py       # AST unit tests
├── test_parser.py       # Parser integration tests
└── test_renderer.py     # Renderer unit tests
```

## Development

```bash
# Clone and set up
git clone <repo>
cd docx2html
python -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"

# Run tests
pytest

# Run tests with coverage
pytest --cov=docx2html --cov-report=term-missing
```

## Supported Elements

- Headings (levels 1–6)
- Paragraphs with inline bold, italic, underline, strikethrough, monospace, hyperlinks
- Unordered lists (`List Bullet` style + `numPr` detection)
- Ordered lists (`List Number` style + `numPr` detection)
- Nested lists (depth-aware)
- Tables with header row detection
- Images (extracted to `media/`)
- Code blocks (`Code`, `Verbatim`, monospace-font paragraphs)
- Horizontal rules (paragraph border-bottom)
- Document metadata (title, author, created date)

## License

MIT
