"""Citire și comparare de documente, independentă de interfața PyQt.

Funcțiile publice păstrează forma API-ului folosită anterior în ``main.py``.
Opțiunile de comparare afectează numai cheile folosite pentru potrivire; blocurile
returnate păstrează întotdeauna textul original pentru afișare și export.
"""

from __future__ import annotations

import difflib
import html as html_lib
from collections.abc import Callable, Mapping
from pathlib import Path
from typing import Any, Optional

import fitz  # PyMuPDF
from docx import Document as DocxDocument
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph as DocxParagraph

from notifications import UserFacingError


CancellationCallback = Optional[Callable[[], bool]]

DEFAULT_COMPARISON_OPTIONS = {
    "ignore_whitespace": False,
    "ignore_case": False,
    "ignore_headers": False,
}


class ComparisonCancelled(UserFacingError):
    """Ridicată când apelantul solicită anularea citirii sau a comparației."""

    def __init__(self, message: str = "Comparația a fost anulată.") -> None:
        super().__init__(message)


def _check_cancel(cancel: CancellationCallback) -> None:
    """Raise the dedicated cancellation exception when cancellation was requested."""
    if cancel is not None and cancel():
        raise ComparisonCancelled()


def _is_vertical_merge_continuation(cell_element: Any) -> bool:
    """Whether a table cell only continues the cell immediately above it.

    Word normally leaves continuation cells empty, but some producers retain their
    paragraphs. Skipping them prevents the logical merged cell from being emitted
    twice while retaining the content of its ``restart`` cell.
    """
    tc_pr = cell_element.find(qn("w:tcPr"))
    if tc_pr is None:
        return False
    v_merge = tc_pr.find(qn("w:vMerge"))
    if v_merge is None:
        return False
    return v_merge.get(qn("w:val")) != "restart"


def _iter_block_texts(
    element: Any,
    doc: DocxDocument,
    cancel: CancellationCallback = None,
    _seen_cells: Optional[set[Any]] = None,
) -> list[str]:
    """Return non-empty paragraphs in document order from paragraphs and tables.

    The routine walks raw WordprocessingML so body, header/footer, nested-table,
    and merged-cell text use the same ordering. A physical table cell is visited
    once only. Vertical merge continuations are intentionally ignored because
    their text belongs to the preceding logical cell.
    """
    seen_cells = _seen_cells if _seen_cells is not None else set()
    texts: list[str] = []

    for child in element.iterchildren():
        _check_cancel(cancel)
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag == "p":
            text = DocxParagraph(child, doc).text.strip()
            if text:
                texts.append(text)
        elif tag == "tbl":
            for row in child.findall(qn("w:tr")):
                _check_cancel(cancel)
                for cell in row.findall(qn("w:tc")):
                    _check_cancel(cancel)
                    # Keep the actual lxml elements alive in the set. Using
                    # ``id(cell)`` alone is unsafe: an earlier wrapper can be
                    # released during traversal and Python may reuse its ID for
                    # a different physical cell in a later row.
                    if cell in seen_cells:
                        continue
                    seen_cells.add(cell)
                    if _is_vertical_merge_continuation(cell):
                        continue
                    texts.extend(_iter_block_texts(cell, doc, cancel, seen_cells))
    return texts


def _part_identity(part: Any) -> str:
    """Use the package part name when available, falling back to object identity."""
    part_name = getattr(part, "partname", None)
    return str(part_name) if part_name is not None else str(id(part))


def _iter_header_footer_texts(
    doc: DocxDocument,
    cancel: CancellationCallback = None,
) -> list[str]:
    """Extract header/footer paragraphs and tables, once per part and role.

    Identical *text* is deliberately not a de-duplication key: a header and a
    footer may legitimately contain the same words, as may two distinct section
    parts. Linked sections are skipped because their inherited part is emitted at
    the section where it is defined. Prefixes remain compatible with the former
    UI/API: ``[HEADER]`` and ``[FOOTER]``.
    """
    texts: list[str] = []
    seen_parts: set[tuple[str, str]] = set()
    roles = (
        ("HEADER", "header", "header"),
        ("HEADER", "first_page_header", "header"),
        ("HEADER", "even_page_header", "header"),
        ("FOOTER", "footer", "footer"),
        ("FOOTER", "first_page_footer", "footer"),
        ("FOOTER", "even_page_footer", "footer"),
    )

    for section in doc.sections:
        _check_cancel(cancel)
        for tag, attribute, role in roles:
            _check_cancel(cancel)
            story = getattr(section, attribute)
            # Accessing ``part`` on an inherited/undefined story can create a new
            # empty header/footer part, so check linkage first.
            if story is None or story.is_linked_to_previous:
                continue
            key = (role, _part_identity(story.part))
            if key in seen_parts:
                continue
            seen_parts.add(key)
            for text in _iter_block_texts(story._element, doc, cancel):
                _check_cancel(cancel)
                texts.append(f"[{tag}] {text}")
    return texts


def read_docx(path: str, cancel: CancellationCallback = None) -> list[str]:
    """Read body plus defined headers and footers from a DOCX file."""
    _check_cancel(cancel)
    doc = DocxDocument(path)
    _check_cancel(cancel)
    header_footer_texts = _iter_header_footer_texts(doc, cancel)
    body_texts = _iter_block_texts(doc.element.body, doc, cancel)
    _check_cancel(cancel)
    return header_footer_texts + body_texts


def _pdf_empty_error() -> UserFacingError:
    return UserFacingError(
        "Fișierul PDF este gol și nu conține pagini care pot fi comparate."
    )


def _pdf_image_only_error() -> UserFacingError:
    return UserFacingError(
        "Nu s-a putut extrage text din PDF. Documentul pare scanat ca imagine "
        "sau nu conține text selectabil; aplicația nu efectuează OCR."
    )


def read_pdf(path: str, cancel: CancellationCallback = None) -> list[str]:
    """Extract text blocks from a text-based, unencrypted PDF.

    The application has no OCR engine. PDFs that contain only scans/images and
    encrypted PDFs are reported explicitly rather than silently returning an
    empty comparison input.
    """
    _check_cancel(cancel)
    try:
        document = fitz.open(path)
    except fitz.EmptyFileError as error:
        raise _pdf_empty_error() from error

    with document:
        _check_cancel(cancel)
        if document.needs_pass:
            raise UserFacingError(
                "Documentul PDF este protejat prin parolă și nu poate fi comparat. "
                "Eliminați parola sau selectați o versiune neprotejată."
            )
        if document.page_count == 0:
            raise _pdf_empty_error()

        paragraphs: list[str] = []
        has_selectable_text = False
        for page in document:
            _check_cancel(cancel)
            # This separate check distinguishes a genuinely image-only PDF from
            # a document whose text blocks merely do not meet display filtering.
            if page.get_text().strip():
                has_selectable_text = True
            blocks = page.get_text("blocks")
            blocks_sorted = sorted(blocks, key=lambda block: (round(block[1] / 12) * 12, block[0]))
            for block in blocks_sorted:
                _check_cancel(cancel)
                if len(block) > 6 and block[6] == 0:
                    text = " ".join(block[4].split())
                    if text:
                        paragraphs.append(text)

        if not has_selectable_text or not paragraphs:
            raise _pdf_image_only_error()
        return paragraphs


def read_document(path: str, cancel: CancellationCallback = None) -> list[str]:
    """Read a supported document, dispatching on the case-insensitive suffix."""
    _check_cancel(cancel)
    extension = Path(path).suffix.lower()
    if extension == ".docx":
        return read_docx(path, cancel)
    if extension == ".pdf":
        return read_pdf(path, cancel)
    raise UserFacingError(
        f"Format neacceptat: {extension}\nSunt acceptate fișierele .docx și .pdf."
    )


def _resolved_options(options: Optional[Mapping[str, Any]]) -> dict[str, bool]:
    """Return the supported comparison flags without mutating caller input."""
    if options is None:
        return dict(DEFAULT_COMPARISON_OPTIONS)
    if not isinstance(options, Mapping):
        raise TypeError("Opțiunile de comparație trebuie transmise într-un dicționar.")
    return {
        name: bool(options.get(name, default))
        for name, default in DEFAULT_COMPARISON_OPTIONS.items()
    }


def _is_header_footer(text: str) -> bool:
    return text.startswith("[HEADER] ") or text.startswith("[FOOTER] ")


def _normalize_for_comparison(text: str, options: Mapping[str, bool]) -> str:
    """Produce a match key while leaving caller-owned/display text untouched."""
    normalized = text
    if options["ignore_whitespace"]:
        # Collapse formatting whitespace but retain word boundaries (unlike
        # removing all whitespace, which would make "nu mai" equal "numai").
        normalized = " ".join(normalized.split())
    if options["ignore_case"]:
        normalized = normalized.casefold()
    return normalized


def word_level_diff(
    text1: str,
    text2: str,
    cancel: CancellationCallback = None,
) -> tuple[str, str]:
    """Return safe HTML snippets highlighting word-level differences."""
    _check_cancel(cancel)
    words1 = text1.split()
    words2 = text2.split()
    matcher = difflib.SequenceMatcher(None, words1, words2, autojunk=False)

    left_parts: list[str] = []
    right_parts: list[str] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        _check_cancel(cancel)
        left_words = html_lib.escape(" ".join(words1[i1:i2]))
        right_words = html_lib.escape(" ".join(words2[j1:j2]))

        if tag == "equal":
            left_parts.append(left_words)
            right_parts.append(left_words)
        elif tag == "replace":
            left_parts.append(
                f'<span style="background:#FFCDD2;text-decoration:line-through;'
                f'border-radius:2px;padding:0 2px;">{left_words}</span>'
            )
            right_parts.append(
                f'<span style="background:#C8E6C9;border-radius:2px;padding:0 2px;">{right_words}</span>'
            )
        elif tag == "delete":
            left_parts.append(
                f'<span style="background:#FFCDD2;text-decoration:line-through;'
                f'border-radius:2px;padding:0 2px;">{left_words}</span>'
            )
        elif tag == "insert":
            right_parts.append(
                f'<span style="background:#C8E6C9;border-radius:2px;padding:0 2px;">{right_words}</span>'
            )

    return " ".join(left_parts), " ".join(right_parts)


def compare_documents(
    paras1: list[str],
    paras2: list[str],
    options: Optional[Mapping[str, Any]] = None,
    cancel: CancellationCallback = None,
) -> dict[str, Any]:
    """Compare document texts while retaining originals in result blocks.

    ``options`` may contain these boolean flags:

    * ``ignore_whitespace`` — collapse whitespace for matching;
    * ``ignore_case`` — compare with Unicode-aware ``casefold``;
    * ``ignore_headers`` — exclude ``[HEADER]``/``[FOOTER]`` rows from the
      comparison entirely.

    The return shape remains compatible with the existing render/export code and
    additionally contains the resolved ``options`` dictionary.
    """
    resolved = _resolved_options(options)
    _check_cancel(cancel)

    left_original: list[str] = []
    for paragraph in paras1:
        _check_cancel(cancel)
        if not resolved["ignore_headers"] or not _is_header_footer(paragraph):
            left_original.append(paragraph)

    right_original: list[str] = []
    for paragraph in paras2:
        _check_cancel(cancel)
        if not resolved["ignore_headers"] or not _is_header_footer(paragraph):
            right_original.append(paragraph)

    left_keys: list[str] = []
    for paragraph in left_original:
        _check_cancel(cancel)
        left_keys.append(_normalize_for_comparison(paragraph, resolved))

    right_keys: list[str] = []
    for paragraph in right_original:
        _check_cancel(cancel)
        right_keys.append(_normalize_for_comparison(paragraph, resolved))

    _check_cancel(cancel)
    matcher = difflib.SequenceMatcher(None, left_keys, right_keys, autojunk=False)
    similarity = matcher.ratio()
    _check_cancel(cancel)

    left_blocks: list[tuple[str, str, bool]] = []
    right_blocks: list[tuple[str, str, bool]] = []
    stats = {
        "total_left": len(left_original),
        "total_right": len(right_original),
        "equal": 0,
        "modified": 0,
        "added": 0,
        "deleted": 0,
    }

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        _check_cancel(cancel)
        if tag == "equal":
            for offset in range(i2 - i1):
                _check_cancel(cancel)
                left_blocks.append(("equal", left_original[i1 + offset], False))
                right_blocks.append(("equal", right_original[j1 + offset], False))
                stats["equal"] += 1
        elif tag == "replace":
            left_slice = left_original[i1:i2]
            right_slice = right_original[j1:j2]
            common = min(len(left_slice), len(right_slice))
            stats["modified"] += common
            stats["deleted"] += len(left_slice) - common
            stats["added"] += len(right_slice) - common

            for offset in range(max(len(left_slice), len(right_slice))):
                _check_cancel(cancel)
                has_left = offset < len(left_slice)
                has_right = offset < len(right_slice)
                if has_left and has_right:
                    left_html, right_html = word_level_diff(
                        left_slice[offset], right_slice[offset], cancel
                    )
                    left_blocks.append(("replace", left_html, True))
                    right_blocks.append(("replace", right_html, True))
                elif has_left:
                    left_blocks.append(("delete", left_slice[offset], False))
                    right_blocks.append(("placeholder", "", False))
                else:
                    left_blocks.append(("placeholder", "", False))
                    right_blocks.append(("insert", right_slice[offset], False))
        elif tag == "delete":
            for offset in range(i2 - i1):
                _check_cancel(cancel)
                left_blocks.append(("delete", left_original[i1 + offset], False))
                right_blocks.append(("placeholder", "", False))
                stats["deleted"] += 1
        elif tag == "insert":
            for offset in range(j2 - j1):
                _check_cancel(cancel)
                left_blocks.append(("placeholder", "", False))
                right_blocks.append(("insert", right_original[j1 + offset], False))
                stats["added"] += 1

    diff_positions: list[int] = []
    in_difference = False
    for index, (left_block, right_block) in enumerate(zip(left_blocks, right_blocks)):
        _check_cancel(cancel)
        left_status = left_block[0]
        right_status = right_block[0]
        is_difference = (
            left_status not in ("equal", "placeholder")
            or right_status not in ("equal", "placeholder")
        )
        if is_difference and not in_difference:
            diff_positions.append(index)
        in_difference = is_difference

    _check_cancel(cancel)
    return {
        "similarity": similarity,
        "left_blocks": left_blocks,
        "right_blocks": right_blocks,
        "diff_positions": diff_positions,
        "stats": stats,
        "diff_count": len(diff_positions),
        "options": resolved,
    }
