"""
utils.py
--------
Shared utility functions used across the docx2html package.

All functions here are pure (no side effects) and have no dependencies on
other docx2html modules, making them safe to import anywhere in the package.
"""

from __future__ import annotations

import json
import logging
import sys
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Path helpers
# ---------------------------------------------------------------------------


def resolve_output_path(input_path: Path, output_path: str | Path | None) -> Path:
    """
    Determine the output HTML file path.

    Rules
    ~~~~~
    1. If *output_path* is given explicitly, use it as-is.
    2. Otherwise, replace the ``.docx`` extension with ``.html`` in the same
       directory as *input_path*.

    Parameters
    ----------
    input_path:
        Path to the source ``.docx`` file.
    output_path:
        Explicit output path, or ``None`` to derive from *input_path*.

    Returns
    -------
    Path
        Resolved output path.
    """
    if output_path is not None:
        return Path(output_path)
    return input_path.with_suffix(".html")


def resolve_media_dir(output_path: Path, media_dir: str | Path | None = None) -> Path:
    """
    Determine where to save extracted images.

    Defaults to a ``media/`` directory that is a sibling of *output_path*.

    Parameters
    ----------
    output_path:
        The resolved HTML output path.
    media_dir:
        Explicit media directory, or ``None`` to derive automatically.

    Returns
    -------
    Path
        Resolved media directory path.
    """
    if media_dir is not None:
        return Path(media_dir)
    return output_path.parent / "media"


# ---------------------------------------------------------------------------
# JSON serialisation
# ---------------------------------------------------------------------------


def to_json(data: Any, *, indent: int | None = 2, sort_keys: bool = False) -> str:
    """
    Serialise *data* to a JSON string.

    Parameters
    ----------
    data:
        Any JSON-serialisable Python object.
    indent:
        Indentation level.  Pass ``None`` for compact output.
    sort_keys:
        Sort dictionary keys alphabetically.

    Returns
    -------
    str
        JSON-encoded string.
    """
    return json.dumps(data, indent=indent, sort_keys=sort_keys, ensure_ascii=False)


def from_json(text: str) -> Any:
    """Deserialise *text* from JSON."""
    return json.loads(text)


# ---------------------------------------------------------------------------
# File I/O helpers
# ---------------------------------------------------------------------------


def write_text(path: Path, content: str, encoding: str = "utf-8") -> None:
    """
    Write *content* to *path*, creating parent directories as needed.

    Parameters
    ----------
    path:
        Target file path.
    content:
        Text to write.
    encoding:
        File encoding (default ``utf-8``).
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding=encoding)
    logger.debug("Wrote %d bytes to %s", len(content.encode(encoding)), path)


def read_text(path: Path, encoding: str = "utf-8") -> str:
    """Read and return the text content of *path*."""
    return path.read_text(encoding=encoding)


# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------


def configure_logging(verbose: bool = False) -> None:
    """
    Configure the root logger for CLI usage.

    Parameters
    ----------
    verbose:
        When ``True``, set the log level to ``DEBUG``.  Otherwise ``WARNING``
        is used so only error-level messages appear in normal operation.
    """
    level = logging.DEBUG if verbose else logging.WARNING
    logging.basicConfig(
        level=level,
        format="%(levelname)s  %(name)s: %(message)s",
        stream=sys.stderr,
    )


# ---------------------------------------------------------------------------
# Validation helpers
# ---------------------------------------------------------------------------


def validate_docx_path(path: Path) -> None:
    """
    Raise a :class:`ValueError` or :class:`FileNotFoundError` if *path* is not
    a valid, readable ``.docx`` file.

    Parameters
    ----------
    path:
        Path to validate.

    Raises
    ------
    FileNotFoundError
        When the file does not exist.
    ValueError
        When the path does not point to a ``.docx`` file.
    """
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")
    if not path.is_file():
        raise ValueError(f"Input path is not a file: {path}")
    if path.suffix.lower() != ".docx":
        raise ValueError(
            f"Expected a .docx file, got {path.suffix!r}.  "
            f"Only Microsoft Word documents are supported."
        )


# ---------------------------------------------------------------------------
# String helpers
# ---------------------------------------------------------------------------


def truncate(text: str, max_len: int = 80, ellipsis: str = "…") -> str:
    """Return *text* truncated to *max_len* characters, appending *ellipsis* if cut."""
    if len(text) <= max_len:
        return text
    return text[: max_len - len(ellipsis)] + ellipsis


def slugify(text: str) -> str:
    """
    Convert *text* to a URL-safe slug.

    Example::

        >>> slugify("Hello, World!")
        'hello-world'
    """
    import re
    text = text.lower().strip()
    text = re.sub(r"[^\w\s-]", "", text)
    text = re.sub(r"[\s_-]+", "-", text)
    text = re.sub(r"^-+|-+$", "", text)
    return text
