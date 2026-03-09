"""
cli.py
------
Command-line interface for docx2html.

Usage examples
~~~~~~~~~~~~~~
::

    # Convert to HTML (output inferred from input filename)
    docx2html template.docx

    # Explicit output path
    docx2html template.docx -o output.html

    # Print JSON AST to stdout
    docx2html template.docx --json

    # Pretty-printed HTML to stdout (no file written)
    docx2html template.docx --stdout

    # Compact (non-pretty) HTML
    docx2html template.docx --no-pretty

    # Verbose logging (useful for debugging)
    docx2html template.docx -v

    # Fragment only (no <!DOCTYPE html> wrapper)
    docx2html template.docx --fragment

The CLI is built with `typer` for ergonomic argument parsing and
`rich` for coloured terminal output.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.panel import Panel
from rich.syntax import Syntax
from rich.text import Text

from docx2html import __version__
from docx2html import parser as _parser
from docx2html import renderer_html as _renderer
from docx2html.utils import (
    configure_logging,
    resolve_media_dir,
    resolve_output_path,
    to_json,
    validate_docx_path,
    write_text,
)

# ---------------------------------------------------------------------------
# Typer app
# ---------------------------------------------------------------------------

app = typer.Typer(
    name="docx2html",
    help=(
        "Convert Microsoft Word (.docx) files to clean, structured HTML "
        "optimised for LLM ingestion and editing."
    ),
    add_completion=False,
    pretty_exceptions_enable=False,
)

_console = Console(stderr=True)   # status / error messages → stderr
_out_console = Console()           # final output → stdout


# ---------------------------------------------------------------------------
# Version callback
# ---------------------------------------------------------------------------


def _version_callback(value: bool) -> None:
    if value:
        typer.echo(f"docx2html {__version__}")
        raise typer.Exit()


# ---------------------------------------------------------------------------
# Main command
# ---------------------------------------------------------------------------


@app.command()
def convert(
    input_file: Path = typer.Argument(
        ...,
        help="Path to the source .docx file.",
        exists=True,
        readable=True,
        resolve_path=True,
    ),
    output: Optional[Path] = typer.Option(
        None,
        "--output",
        "-o",
        help=(
            "Output HTML file path.  Defaults to the input filename with a "
            ".html extension in the same directory."
        ),
        writable=True,
        resolve_path=True,
    ),
    json_output: bool = typer.Option(
        False,
        "--json",
        help="Print the parsed Document AST as JSON instead of HTML.",
        is_flag=True,
    ),
    pretty: bool = typer.Option(
        True,
        "--pretty/--no-pretty",
        help="Pretty-print the HTML output (enabled by default).",
    ),
    stdout: bool = typer.Option(
        False,
        "--stdout",
        help="Print output to stdout instead of writing to a file.",
        is_flag=True,
    ),
    fragment: bool = typer.Option(
        False,
        "--fragment",
        help=(
            "Emit an HTML fragment (no <!DOCTYPE html> wrapper).  "
            "Useful for embedding in an existing page."
        ),
        is_flag=True,
    ),
    media_dir: Optional[Path] = typer.Option(
        None,
        "--media-dir",
        help=(
            "Directory where extracted images are saved.  "
            "Defaults to a media/ directory next to the output file."
        ),
        resolve_path=True,
    ),
    title: Optional[str] = typer.Option(
        None,
        "--title",
        "-t",
        help="Override the <title> tag in the generated HTML page.",
    ),
    verbose: bool = typer.Option(
        False,
        "--verbose",
        "-v",
        help="Enable debug logging to stderr.",
        is_flag=True,
    ),
    version: Optional[bool] = typer.Option(  # noqa: UP007
        None,
        "--version",
        callback=_version_callback,
        is_eager=True,
        help="Show the version and exit.",
    ),
) -> None:
    """
    Convert *INPUT_FILE* (.docx) to clean, Tailwind-styled HTML.

    By default the output is written to a file with the same stem as the input
    and a .html extension.  Use --stdout to print to stdout, or --json to
    dump the parsed Document AST instead of HTML.
    """
    configure_logging(verbose)

    # ------------------------------------------------------------------
    # Validate input
    # ------------------------------------------------------------------
    try:
        validate_docx_path(input_file)
    except (FileNotFoundError, ValueError) as exc:
        _console.print(f"[bold red]Error:[/bold red] {exc}")
        raise typer.Exit(code=1) from exc

    # ------------------------------------------------------------------
    # Resolve output path
    # ------------------------------------------------------------------
    out_path = resolve_output_path(input_file, output)
    resolved_media_dir = resolve_media_dir(out_path, media_dir)

    # ------------------------------------------------------------------
    # Parse
    # ------------------------------------------------------------------
    _console.print(f"[dim]Parsing[/dim] {input_file.name} …")
    try:
        doc = _parser.parse(input_file, media_dir=resolved_media_dir)
    except Exception as exc:
        _console.print(f"[bold red]Parse error:[/bold red] {exc}")
        if verbose:
            import traceback
            _console.print(traceback.format_exc())
        raise typer.Exit(code=1) from exc

    block_count = len(doc.blocks)
    _console.print(
        f"[dim]Parsed[/dim] [bold]{block_count}[/bold] block(s) "
        f"from [bold]{input_file.name}[/bold]"
    )

    # ------------------------------------------------------------------
    # JSON output mode
    # ------------------------------------------------------------------
    if json_output:
        ast_dict = doc.to_dict()
        json_str = to_json(ast_dict, indent=2 if pretty else None)

        if stdout:
            _out_console.print(
                Syntax(json_str, "json", theme="monokai", line_numbers=False)
            )
        else:
            json_path = out_path.with_suffix(".json")
            try:
                write_text(json_path, json_str)
            except OSError as exc:
                _console.print(f"[bold red]Write error:[/bold red] {exc}")
                raise typer.Exit(code=1) from exc
            _console.print(
                Panel(
                    Text(str(json_path), style="bold green"),
                    title="JSON AST written",
                    expand=False,
                )
            )
        return

    # ------------------------------------------------------------------
    # HTML rendering
    # ------------------------------------------------------------------
    _console.print("[dim]Rendering HTML …[/dim]")
    try:
        html = _renderer.render(
            doc,
            pretty=pretty,
            include_document_wrapper=not fragment,
            title=title,
        )
    except Exception as exc:
        _console.print(f"[bold red]Render error:[/bold red] {exc}")
        if verbose:
            import traceback
            _console.print(traceback.format_exc())
        raise typer.Exit(code=1) from exc

    # ------------------------------------------------------------------
    # Output
    # ------------------------------------------------------------------
    if stdout:
        if pretty:
            _out_console.print(
                Syntax(html, "html", theme="monokai", line_numbers=False)
            )
        else:
            typer.echo(html)
        return

    try:
        write_text(out_path, html)
    except OSError as exc:
        _console.print(f"[bold red]Write error:[/bold red] {exc}")
        raise typer.Exit(code=1) from exc

    size_kb = len(html.encode()) / 1024
    _console.print(
        Panel(
            Text.assemble(
                (str(out_path), "bold green"),
                ("  ", ""),
                (f"({size_kb:.1f} KB)", "dim"),
            ),
            title="[bold]HTML written[/bold]",
            expand=False,
        )
    )

    if doc.blocks and any(
        hasattr(b, "src") for b in doc.blocks  # Image blocks
    ):
        _console.print(f"[dim]Images saved to:[/dim] {resolved_media_dir}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    """Programmatic entry point (also used by ``pyproject.toml`` scripts)."""
    app()


if __name__ == "__main__":
    main()
