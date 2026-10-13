"""Portable HTML and PDF comparison-report exporters.

The public functions intentionally retain the historical positional arguments and
add optional ``progress`` and ``cancelled`` callbacks for use by worker threads.
``progress`` may be a callable or a Qt-style object exposing ``emit``.
"""
from __future__ import annotations

from datetime import datetime
import html as html_lib
from pathlib import Path
import re
import sys
from typing import Any, Callable, Optional

from fpdf import FPDF


class ExportCancelled(RuntimeError):
    """Raised when a cooperative HTML or PDF export is cancelled."""

    def __init__(self, message: str = "Exportul raportului a fost anulat.") -> None:
        super().__init__(message)


_FONT_FAMILY = "ComparerDejaVu"
_TAG_RE = re.compile(r"<[^>]+>")
_HTML_STYLES = {
    "equal": ("#FFFFFF", "#E0E0E0", "3px solid"),
    "replace": ("#FFECB3", "#FFB300", "3px solid"),
    "delete": ("#FFCDD2", "#E53935", "3px solid"),
    "insert": ("#C8E6C9", "#43A047", "3px solid"),
    "placeholder": ("#F9F9F9", "#EEEEEE", "3px dotted"),
}
_HTML_BADGES = {
    "delete": (
        '<span style="font-size:9px;background:#E53935;color:#fff;'
        'padding:1px 6px;border-radius:2px;margin-right:6px;">ȘTERS</span>'
    ),
    "insert": (
        '<span style="font-size:9px;background:#43A047;color:#fff;'
        'padding:1px 6px;border-radius:2px;margin-right:6px;">NOU</span>'
    ),
}


# These public APIs have the same original required positional parameters.
ProgressCallback = Optional[Callable[[str], Any]]
CancelCallback = Optional[Callable[[], bool]]


def _resource(relative: str) -> Path:
    """Return a source-tree or PyInstaller one-file resource path."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative


def _emit_progress(progress: Any, message: str) -> None:
    if progress is None:
        return
    emitter = getattr(progress, "emit", None)
    if callable(emitter):
        emitter(message)
    elif callable(progress):
        progress(message)


def _is_cancelled(cancelled: Any) -> bool:
    if cancelled is None:
        return False
    if callable(cancelled):
        return bool(cancelled())
    is_set = getattr(cancelled, "is_set", None)
    if callable(is_set):
        return bool(is_set())
    return bool(cancelled)


def _check_cancelled(cancelled: Any) -> None:
    if _is_cancelled(cancelled):
        raise ExportCancelled()


def _strip_html_tags(text: str) -> str:
    """Convert trusted word-diff markup back to its printable text."""
    return html_lib.unescape(_TAG_RE.sub("", text)).strip()


def _hex_to_rgb(color: str) -> tuple[int, int, int]:
    value = color.lstrip("#")
    return tuple(int(value[index : index + 2], 16) for index in (0, 2, 4))


def _comparison_options(result: dict) -> dict[str, bool]:
    """Accept current and future result option layouts without changing callers."""
    nested = result.get("comparison_options") or result.get("options") or {}
    if not isinstance(nested, dict):
        nested = {}
    return {
        name: bool(result.get(name, nested.get(name, False)))
        for name in ("ignore_case", "ignore_whitespace", "ignore_headers")
    }


def _options_html(result: dict) -> str:
    options = _comparison_options(result)
    rows = (
        ("Ignoră majusculele/minusculele", options["ignore_case"]),
        ("Ignoră spațiile albe", options["ignore_whitespace"]),
        ("Ignoră anteturile și subsolurile", options["ignore_headers"]),
    )
    return "".join(
        f'<li><span>{label}</span><strong>{"Da" if enabled else "Nu"}</strong></li>'
        for label, enabled in rows
    )


def _score_explanation(result: dict) -> str:
    enabled = [
        label
        for key, label in (
            ("ignore_case", "diferențierea majusculelor/minusculelor"),
            ("ignore_whitespace", "spațiile albe"),
            ("ignore_headers", "anteturile și subsolurile"),
        )
        if _comparison_options(result)[key]
    ]
    if enabled:
        applied = ", ".join(enabled)
        return (
            "Scorul este calculat din secvențele de paragrafe identice identificate "
            f"după normalizare; au fost aplicate: {applied}. Paragrafele modificate "
            "nu sunt numărate ca identice."
        )
    return (
        "Scorul este calculat din secvențele de paragrafe identice. Nu au fost "
        "aplicate opțiuni de normalizare, iar paragrafele modificate nu sunt "
        "numărate ca identice."
    )


def _options_pdf(result: dict) -> str:
    """Return the comparison settings in a compact first-page PDF form."""
    options = _comparison_options(result)
    return " | ".join((
        f"Majuscule/minuscule: {'ignorate' if options['ignore_case'] else 'luate în calcul'}",
        f"Spații albe: {'ignorate' if options['ignore_whitespace'] else 'luate în calcul'}",
        f"Antet/subsol: {'ignorate' if options['ignore_headers'] else 'luate în calcul'}",
    ))


def _should_report(index: int, total: int, interval: int = 50) -> bool:
    """Keep worker-queue progress messages useful without flooding the GUI."""
    return index == total or (index > 0 and index % interval == 0)


def _render_html_block(
    block: tuple,
    side: str,
    block_index: int,
    anchor_map: dict[int, int],
    padding_y: int,
) -> str:
    """Render one static comparison cell, retaining trusted word-diff markup."""
    status, content, is_html = block
    background, border_color, border_width = _HTML_STYLES.get(status, _HTML_STYLES["equal"])
    anchor = f'<a name="diff-{anchor_map[block_index]}"></a>' if block_index in anchor_map else ""
    badge = _HTML_BADGES["delete"] if status == "delete" and side == "left" else ""
    if status == "insert" and side == "right":
        badge = _HTML_BADGES["insert"]
    if status == "placeholder":
        inner = ""
    elif is_html:
        # word_level_diff has already escaped document text; preserving this
        # markup keeps replacement/deletion highlights in the HTML report.
        inner = f"{badge}<span>{content}</span>"
    else:
        inner = f"{badge}<span>{html_lib.escape(str(content))}</span>"
    return (
        f'<div class="report-block" style="background:{background};'
        f'border-left:{border_width} {border_color};padding:{padding_y}px 10px;'
        f'border-radius:0 2px 2px 0;min-height:22px;overflow:hidden;">'
        f"{anchor}{inner}</div>"
    )


def render_html(
    blocks: list,
    side: str,
    diff_positions: Optional[list] = None,
    padding_y: int = 7,
    interactive: bool = False,
    progress: ProgressCallback = None,
    cancelled: CancelCallback = None,
) -> str:
    """Render one report column, retaining the existing trusted word highlights.

    ``interactive`` is preserved for compatibility with existing callers.  The
    output contains anchors for ``diff_positions`` exactly as the original
    exporter did.
    """
    del interactive  # The static export has no interactive behaviour.
    anchor_map = {block_index: group_index for group_index, block_index in enumerate(diff_positions or [])}
    lines = [
        "<!DOCTYPE html><html>",
        '<head><meta charset="utf-8"></head>',
        '<body style="font-family:system-ui,Arial,sans-serif;font-size:13px;'
        'line-height:1.55;margin:0;padding:6px;background:#FAFAFA;color:#111;">',
    ]

    total = len(blocks)
    document_name = "documentul 1" if side == "left" else "documentul 2"
    _emit_progress(progress, f"Se pregătește {document_name} pentru raportul HTML...")
    for block_index, block in enumerate(blocks):
        _check_cancelled(cancelled)
        lines.append(_render_html_block(block, side, block_index, anchor_map, padding_y))
        if _should_report(block_index + 1, total):
            _emit_progress(progress, f"Se pregătește {document_name}: {block_index + 1}/{total} rânduri.")

    _check_cancelled(cancelled)
    lines.append("</body></html>")
    return "\n".join(lines)


def _render_html_pairs(result: dict, progress: ProgressCallback, cancelled: CancelCallback) -> str:
    """Render corresponding blocks into one grid row, keeping both sides aligned."""
    left_blocks = result["left_blocks"]
    right_blocks = result["right_blocks"]
    if len(left_blocks) != len(right_blocks):
        raise ValueError("Blocurile documentelor trebuie să aibă același număr de rânduri.")
    anchor_map = {block_index: group_index for group_index, block_index in enumerate(result.get("diff_positions") or [])}
    total = len(left_blocks)
    lines = ['<div class="paired-report">']
    _emit_progress(progress, "Se pregătesc rândurile pereche ale raportului HTML...")
    for index, (left, right) in enumerate(zip(left_blocks, right_blocks)):
        _check_cancelled(cancelled)
        lines.append('<div class="pair-row">')
        lines.append(_render_html_block(left, "left", index, anchor_map, 7))
        lines.append(_render_html_block(right, "right", index, anchor_map, 7))
        lines.append("</div>")
        if _should_report(index + 1, total):
            _emit_progress(progress, f"Se pregătesc rândurile HTML: {index + 1}/{total}.")
    _check_cancelled(cancelled)
    lines.append("</div>")
    return "\n".join(lines)


def export_html_report(
    path1: str,
    path2: str,
    result: dict,
    progress: ProgressCallback = None,
    cancelled: CancelCallback = None,
) -> str:
    """Return a complete, UTF-8 HTML comparison report in Romanian."""
    _check_cancelled(cancelled)
    sim = float(result["similarity"])
    stats = result["stats"]
    name1 = html_lib.escape(Path(path1).name)
    name2 = html_lib.escape(Path(path2).name)
    paired_rows = _render_html_pairs(result, progress, cancelled)
    _check_cancelled(cancelled)
    _emit_progress(progress, "Raportul HTML este pregătit.")

    return f"""<!DOCTYPE html>
<html lang="ro">
<head>
<meta charset="utf-8">
<title>Comparație — {name1} vs {name2}</title>
<style>
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  body {{ font-family:system-ui,Arial,sans-serif; background:#fff; color:#111; }}
  .hdr {{ background:#111; color:#fff; padding:14px 24px; }}
  .hdr h1 {{ font-size:16px; font-weight:600; }}
  .hdr p {{ font-size:11px; color:#aaa; margin-top:3px; }}
  .stats {{ background:#F4F4F4; border-bottom:1px solid #ddd; padding:10px 24px; display:flex; gap:28px; align-items:center; flex-wrap:wrap; }}
  .stat .val {{ font-size:20px; font-weight:700; }}
  .stat .lbl {{ font-size:9px; color:#777; letter-spacing:.5px; }}
  .sim {{ font-size:32px; font-weight:700; }}
  .method {{ padding:14px 24px; background:#FCFCFC; border-bottom:1px solid #ddd; font-size:12px; line-height:1.5; }}
  .method h2 {{ font-size:12px; margin-bottom:4px; }}
  .method ul {{ list-style:none; display:flex; gap:12px 28px; flex-wrap:wrap; margin:8px 0; }}
  .method li {{ display:flex; gap:6px; }}
  .method strong {{ color:#333; }}
  .comparison {{ min-width:0; }}
  .column-headings {{ display:grid; grid-template-columns:minmax(0,1fr) minmax(0,1fr); position:sticky; top:0; z-index:2; }}
  .col-hdr {{ background:#222; color:#fff; padding:7px 12px; font-size:11px; font-weight:600; letter-spacing:.3px; overflow-wrap:anywhere; }}
  .paired-report {{ padding:6px; background:#fff; min-height:80vh; }}
  .pair-row {{ display:grid; grid-template-columns:minmax(0,1fr) minmax(0,1fr); gap:6px; align-items:stretch; margin:2px 0; }}
  .report-block {{ min-width:0; overflow-wrap:anywhere; }}
  @media (max-width:720px) {{ .column-headings, .pair-row {{ grid-template-columns:1fr; }} .column-headings {{ position:static; }} }}
</style>
</head>
<body>
<div class="hdr">
  <h1>Raport comparație documente</h1>
  <p>{name1} &nbsp;&#x27F7;&nbsp; {name2}</p>
</div>
<div class="stats">
  <div class="stat"><div class="sim">{sim * 100:.1f}%</div><div class="lbl">SIMILARITATE</div></div>
  <div style="width:1px;background:#ccc;height:36px;"></div>
  <div class="stat"><div class="val">{stats["equal"]}</div><div class="lbl">IDENTICE</div></div>
  <div class="stat"><div class="val">{stats["modified"]}</div><div class="lbl">MODIFICATE</div></div>
  <div class="stat"><div class="val">{stats["added"]}</div><div class="lbl">ADĂUGATE</div></div>
  <div class="stat"><div class="val">{stats["deleted"]}</div><div class="lbl">ȘTERSE</div></div>
  <div class="stat"><div class="val">{stats["total_left"]}</div><div class="lbl">PARAGRAFE DOC. 1</div></div>
  <div class="stat"><div class="val">{stats["total_right"]}</div><div class="lbl">PARAGRAFE DOC. 2</div></div>
</div>
<section class="method" aria-label="Metodă de comparație">
  <h2>Opțiuni de comparație</h2>
  <ul>{_options_html(result)}</ul>
  <p>{_score_explanation(result)}</p>
</section>
<div class="comparison">
  <div class="column-headings">
    <div class="col-hdr">Document 1 — Bază: {name1}</div>
    <div class="col-hdr">Document 2 — Comparație: {name2}</div>
  </div>
  {paired_rows}
</div>
</body></html>"""


def _register_bundled_fonts(pdf: FPDF) -> None:
    """Register app-bundled DejaVu fonts; never probe system font directories."""
    regular = _resource("assets/fonts/DejaVuSans.ttf")
    bold = _resource("assets/fonts/DejaVuSans-Bold.ttf")
    missing = [str(path) for path in (regular, bold) if not path.is_file()]
    if missing:
        raise FileNotFoundError("Lipsesc fonturile incluse pentru export PDF: " + ", ".join(missing))
    pdf.add_font(_FONT_FAMILY, "", str(regular))
    pdf.add_font(_FONT_FAMILY, "B", str(bold))


def _pdf_text(status: str, content: Any, is_html: bool) -> str:
    text = _strip_html_tags(str(content)) if is_html else str(content)
    if status == "placeholder":
        return ""
    if status == "delete":
        return "[ȘTERS] " + text
    if status == "insert":
        return "[NOU] " + text
    return text


def _wrapped_lines(pdf: FPDF, text: str, width: float, line_height: float) -> list[str]:
    """Use fpdf2's measured dry run as the sole line-break source for drawing."""
    if not text:
        return []
    pdf.set_font(_FONT_FAMILY, "", 8)
    return list(pdf.multi_cell(width, line_height, text, dry_run=True, output="LINES"))


def _ellipsize(pdf: FPDF, text: str, width: float) -> str:
    """Fit a filename inside one PDF column without drawing into its neighbour."""
    if pdf.get_string_width(text) <= width:
        return text
    ellipsis = "…"
    if pdf.get_string_width(ellipsis) > width:
        return ""
    shortened = text
    while shortened and pdf.get_string_width(shortened + ellipsis) > width:
        shortened = shortened[:-1]
    return shortened + ellipsis


def export_pdf_report(
    path1: str,
    path2: str,
    result: dict,
    output_path: str,
    progress: ProgressCallback = None,
    cancelled: CancelCallback = None,
) -> None:
    """Write a portable, paired-column PDF report with manual pagination.

    Each logical comparison row is split into equally tall line segments.  Both
    sides advance by the same segment length, so a long paragraph on either
    side cannot cause duplicate, omitted, or vertically desynchronised content.
    Table headers are emitted on the first and every subsequent PDF page.
    """
    _check_cancelled(cancelled)
    sim = float(result["similarity"])
    stats = result["stats"]
    name1 = Path(path1).name
    name2 = Path(path2).name
    left_blocks = result["left_blocks"]
    right_blocks = result["right_blocks"]
    if len(left_blocks) != len(right_blocks):
        raise ValueError("Blocurile documentelor trebuie să aibă același număr de rânduri.")

    backgrounds = {
        "equal": "#FFFFFF", "replace": "#FFECB3", "delete": "#FFCDD2",
        "insert": "#C8E6C9", "placeholder": "#F5F5F5",
    }
    borders = {
        "equal": "#E0E0E0", "replace": "#FFB300", "delete": "#E53935",
        "insert": "#43A047", "placeholder": "#EEEEEE",
    }
    margin_x, margin_top, margin_bottom = 10.0, 10.0, 12.0
    line_height, cell_padding = 3.8, 1.5

    pdf = FPDF(unit="mm", format="A4")
    pdf.set_margins(margin_x, margin_top, margin_x)
    # Page changes are controlled below.  This prevents FPDF from splitting a
    # multi_cell independently of the matching column.
    pdf.set_auto_page_break(False)
    _register_bundled_fonts(pdf)
    content_width = pdf.w - (margin_x * 2)
    column_width = content_width / 2
    text_width = column_width - 4
    body_bottom = pdf.h - margin_bottom
    page_number = 0
    _emit_progress(progress, "Se pregătește raportul PDF...")

    def table_header() -> None:
        pdf.set_fill_color(34, 34, 34)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font(_FONT_FAMILY, "B", 8)
        left_prefix = "  Document 1 — Bază: "
        right_prefix = "  Document 2 — Comparație: "
        left_heading = left_prefix + _ellipsize(pdf, name1, column_width - pdf.get_string_width(left_prefix) - 2)
        right_heading = right_prefix + _ellipsize(pdf, name2, column_width - pdf.get_string_width(right_prefix) - 2)
        pdf.set_x(margin_x)
        pdf.cell(column_width, 7, left_heading, fill=True)
        pdf.cell(column_width, 7, right_heading, fill=True, new_x="LMARGIN", new_y="NEXT")
        pdf.set_text_color(0, 0, 0)

    def add_page(first: bool = False) -> None:
        nonlocal page_number
        _check_cancelled(cancelled)
        pdf.add_page()
        page_number += 1
        if first:
            generated = datetime.now().strftime("%d.%m.%Y  %H:%M:%S")
            pdf.set_font(_FONT_FAMILY, "B", 13)
            pdf.set_text_color(17, 17, 17)
            pdf.cell(content_width, 8, "Raport comparație documente", align="C", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font(_FONT_FAMILY, "", 8)
            pdf.set_text_color(60, 60, 60)
            left_file = "Document 1: "
            right_file = "Document 2: "
            pdf.cell(column_width, 4, left_file + _ellipsize(pdf, name1, column_width - pdf.get_string_width(left_file) - 1))
            pdf.cell(column_width, 4, right_file + _ellipsize(pdf, name2, column_width - pdf.get_string_width(right_file) - 1), new_x="LMARGIN", new_y="NEXT")
            pdf.set_font(_FONT_FAMILY, "", 9)
            pdf.set_text_color(100, 100, 100)
            pdf.cell(content_width, 5, f"Data raport: {generated}", align="C", new_x="LMARGIN", new_y="NEXT")
            pdf.ln(4)
            stat_items = (
                (f"{sim * 100:.1f}%", "SIMILARITATE"), (str(stats["equal"]), "IDENTICE"),
                (str(stats["modified"]), "MODIFICATE"), (str(stats["added"]), "ADĂUGATE"),
                (str(stats["deleted"]), "ȘTERSE"), (str(stats["total_left"]), "PARAGRAFE DOC. 1"),
                (str(stats["total_right"]), "PARAGRAFE DOC. 2"),
            )
            stat_width = content_width / len(stat_items)
            stat_y = pdf.get_y()
            pdf.set_font(_FONT_FAMILY, "B", 10)
            pdf.set_text_color(17, 17, 17)
            for index, (value, _) in enumerate(stat_items):
                pdf.set_xy(margin_x + index * stat_width, stat_y)
                pdf.cell(stat_width, 6, value, align="C")
            pdf.set_font(_FONT_FAMILY, "", 5.6)
            pdf.set_text_color(110, 110, 110)
            for index, (_, label) in enumerate(stat_items):
                pdf.set_xy(margin_x + index * stat_width, stat_y + 6)
                pdf.cell(stat_width, 4, label, align="C")
            pdf.set_y(stat_y + 13)
            pdf.ln(1)
            pdf.set_font(_FONT_FAMILY, "B", 7)
            pdf.set_text_color(35, 35, 35)
            pdf.cell(content_width, 4, "Opțiuni de comparație", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font(_FONT_FAMILY, "", 6.5)
            pdf.set_text_color(75, 75, 75)
            pdf.multi_cell(content_width, 3.4, _options_pdf(result), new_x="LMARGIN", new_y="NEXT")
            pdf.multi_cell(content_width, 3.4, _score_explanation(result), new_x="LMARGIN", new_y="NEXT")
            pdf.ln(2)
        table_header()
        if page_number == 1 or page_number % 50 == 0:
            _emit_progress(progress, f"Se pregătește pagina PDF {page_number}.")

    def text_color(status: str, side: str) -> tuple[int, int, int]:
        if status == "delete" and side == "left":
            return 180, 30, 30
        if status == "insert" and side == "right":
            return 30, 120, 30
        if status == "replace":
            return 50, 50, 50
        return 17, 17, 17

    def draw_segment(
        left_status: str,
        right_status: str,
        left_lines: list[str],
        right_lines: list[str],
        segment_lines: int,
    ) -> None:
        y = pdf.get_y()
        row_height = max(6.0, segment_lines * line_height + cell_padding * 2)
        for x, status in ((margin_x, left_status), (margin_x + column_width, right_status)):
            pdf.set_fill_color(*_hex_to_rgb(backgrounds.get(status, backgrounds["equal"])))
            pdf.rect(x, y, column_width, row_height, "F")
            pdf.set_draw_color(*_hex_to_rgb(borders.get(status, borders["equal"])))
            pdf.set_line_width(0.6)
            pdf.line(x, y, x, y + row_height)
        pdf.set_draw_color(220, 220, 220)
        pdf.set_line_width(0.15)
        pdf.line(margin_x, y + row_height, margin_x + content_width, y + row_height)

        for side, status, x, lines in (
            ("left", left_status, margin_x + 2, left_lines),
            ("right", right_status, margin_x + column_width + 2, right_lines),
        ):
            pdf.set_font(_FONT_FAMILY, "", 8)
            pdf.set_text_color(*text_color(status, side))
            for index, line in enumerate(lines):
                pdf.set_xy(x, y + cell_padding + index * line_height)
                # Lines came from multi_cell(..., dry_run=True, output="LINES"),
                # so this cell draw uses exactly the measured wrapping plan.
                pdf.cell(text_width, line_height, line, new_x="RIGHT", new_y="TOP")
        pdf.set_y(y + row_height)

    add_page(first=True)
    total_rows = len(left_blocks)
    for row_index, (left_block, right_block) in enumerate(zip(left_blocks, right_blocks), start=1):
        _check_cancelled(cancelled)
        left_status, left_content, left_is_html = left_block
        right_status, right_content, right_is_html = right_block
        left_lines = _wrapped_lines(pdf, _pdf_text(left_status, left_content, left_is_html), text_width, line_height)
        right_lines = _wrapped_lines(pdf, _pdf_text(right_status, right_content, right_is_html), text_width, line_height)
        position = 0
        line_count = max(len(left_lines), len(right_lines), 1)

        while position < line_count:
            _check_cancelled(cancelled)
            available_lines = int((body_bottom - pdf.get_y() - cell_padding * 2) // line_height)
            if available_lines < 1:
                add_page()
                continue
            segment_count = min(available_lines, line_count - position)
            # Both columns use the identical [position:end] range.  Empty lines
            # reserve their matching space when the opposite paragraph is longer.
            end = position + segment_count
            draw_segment(
                left_status, right_status, left_lines[position:end], right_lines[position:end], segment_count,
            )
            position = end
            if position < line_count:
                add_page()
        if _should_report(row_index, total_rows):
            _emit_progress(progress, f"Se exportă rândurile PDF: {row_index}/{total_rows}.")

    _check_cancelled(cancelled)
    _emit_progress(progress, "Se finalizează raportul PDF...")
    pdf.output(str(output_path))
    _emit_progress(progress, "Raportul PDF a fost salvat.")
