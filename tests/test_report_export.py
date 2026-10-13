import os
import tempfile
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import fitz

import report_export
from report_export import ExportCancelled, export_html_report, export_pdf_report, render_html


def _result(left_blocks, right_blocks, **extra):
    result = {
        "similarity": 0.625,
        "left_blocks": left_blocks,
        "right_blocks": right_blocks,
        "diff_positions": [0],
        "stats": {
            "equal": 1,
            "modified": 1,
            "added": 0,
            "deleted": 0,
            "total_left": len(left_blocks),
            "total_right": len(right_blocks),
        },
    }
    result.update(extra)
    return result


def _tokens(prefix, count):
    return " ".join(f"{prefix}{number:04d}" for number in range(count))


class ReportExportTests(unittest.TestCase):
    def _pdf_texts(self, path):
        with fitz.open(path) as document:
            return [page.get_text() for page in document]

    def test_html_preserves_trusted_highlights_escapes_filenames_and_explains_options(self):
        highlight = '<span style="background:#FFCDD2">vechi &amp; sigur</span>'
        result = _result(
            [("replace", highlight, True), ("delete", "eliminat", False)],
            [("replace", "nou <sigur>", False), ("placeholder", "", False)],
            ignore_case=True,
            comparison_options={"ignore_whitespace": True, "ignore_headers": True},
        )
        progress = []
        report = export_html_report("/tmp/<bază&>.docx", "/tmp/țintă>.docx", result, progress.append)

        self.assertIn(highlight, report)
        self.assertIn("nou &lt;sigur&gt;", report)
        self.assertIn("&lt;bază&amp;&gt;.docx", report)
        self.assertNotIn("<bază&>.docx", report)
        self.assertIn('class="paired-report"', report)
        self.assertEqual(report.count('class="pair-row"'), 2)
        self.assertIn("position:sticky", report)
        self.assertIn(">ȘTERS<", report)
        self.assertIn("Ignoră majusculele/minusculele", report)
        self.assertEqual(report.count(">Da<"), 3)
        self.assertIn("spațiile albe", report)
        self.assertIn("secvențele de paragrafe identice", report)
        self.assertIn("modificate nu sunt numărate ca identice", report)
        self.assertTrue(progress)

    def test_pdf_long_paragraphs_on_both_sides_keep_all_text_and_repeat_headers(self):
        left_text = _tokens("STANGA", 1200)
        right_text = _tokens("DREAPTA", 1140)
        result = _result([("replace", left_text, False)], [("replace", right_text, False)])
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "ambele-lungi.pdf"
            export_pdf_report("bază.docx", "comparație.docx", result, str(output))
            pages = self._pdf_texts(output)

        self.assertGreater(len(pages), 1)
        joined = "\n".join(pages)
        self.assertIn("Opțiuni de comparație", pages[0])
        self.assertIn("secvențele de paragrafe identice", pages[0])
        for token in ("STANGA0000", "STANGA1199", "DREAPTA0000", "DREAPTA1139"):
            self.assertIn(token, joined)
            self.assertEqual(joined.count(token), 1, token)
        for page in pages:
            self.assertIn("Document 1", page)
            self.assertIn("Document 2", page)

    def test_pdf_one_long_side_and_unicode_are_not_dropped(self):
        left_text = _tokens("UNICSTANGA", 1300) + " Știință țară ÎNGER café Ω Ж"
        right_text = "DREAPTA_SCURTĂ Știință țară ÎNGER café Ω Ж"
        result = _result([("replace", left_text, False)], [("replace", right_text, False)])
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "o-parte-lungă.pdf"
            export_pdf_report("bază.docx", "țintă.docx", result, str(output))
            pages = self._pdf_texts(output)

        self.assertGreater(len(pages), 1)
        joined = "\n".join(pages)
        for value in ("UNICSTANGA0000", "UNICSTANGA1299", "DREAPTA_SCURTĂ", "Știință", "țară", "ÎNGER", "café", "Ω", "Ж"):
            self.assertIn(value, joined)
        source = Path(report_export.__file__).read_text(encoding="utf-8")
        self.assertNotIn("/usr/share/fonts", source)
        self.assertTrue((Path(report_export.__file__).parent / "assets/fonts/DejaVuSans.ttf").is_file())

    def test_cancellation_and_progress_are_cooperative(self):
        blocks = [("equal", f"Rândul {index}", False) for index in range(12)]
        result = _result(blocks, list(blocks))
        html_progress = []
        checks = {"count": 0}

        def cancel_html():
            checks["count"] += 1
            return checks["count"] > 3

        with self.assertRaises(ExportCancelled):
            render_html(blocks, "left", progress=html_progress.append, cancelled=cancel_html)
        self.assertTrue(html_progress)

        pdf_progress = []
        pdf_checks = {"count": 0}

        def cancel_pdf():
            pdf_checks["count"] += 1
            return pdf_checks["count"] > 5

        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "anulat.pdf"
            with self.assertRaises(ExportCancelled):
                export_pdf_report("a.docx", "b.docx", result, str(output), pdf_progress.append, cancel_pdf)
            self.assertFalse(output.exists())
        self.assertTrue(pdf_progress)

    def test_progress_is_throttled_to_roughly_fifty_rows(self):
        blocks = [("equal", f"Rândul {index}", False) for index in range(101)]
        progress = []
        render_html(blocks, "left", progress=progress.append)
        self.assertEqual(len(progress), 4)  # start, 50, 100, final 101
        self.assertTrue(all(isinstance(message, str) for message in progress))


if __name__ == "__main__":
    unittest.main()
