import errno
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import BadZipFile

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import fitz
from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from PyQt5.QtCore import QCoreApplication, QTimer, Qt
from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox
from PyQt5.QtTest import QTest
import time

import main
from notifications import UserFacingError, error_message, show_notification
from romanian_qt import install_romanian_dialogs


class NotificationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        install_romanian_dialogs(cls.app)

    def test_common_errors_are_localized(self):
        cases = [
            (FileNotFoundError("No such file"), "Fișierul nu a fost găsit"),
            (PermissionError("Permission denied"), "Accesul la fișier a fost refuzat"),
            (IsADirectoryError("Is a directory"), "Calea selectată este un dosar"),
            (fitz.EmptyFileError("empty file"), "Fișierul PDF este gol"),
            (fitz.FileDataError("broken document"), "Documentul nu poate fi citit"),
            (BadZipFile("not a zip file"), "Documentul nu poate fi citit"),
            (PackageNotFoundError("Package not found"), "Documentul nu poate fi citit"),
            (
                OSError(errno.ENOSPC, "No space left on device"),
                "Nu există suficient spațiu",
            ),
        ]
        for error, expected in cases:
            with self.subTest(error=type(error).__name__):
                self.assertIn(expected, error_message(error, "Eroare necunoscută."))
                self.assertNotIn(
                    str(error), error_message(error, "Eroare necunoscută.")
                )

    def test_unknown_error_uses_romanian_fallback(self):
        self.assertEqual(
            error_message(RuntimeError("English error"), "Eroare neașteptată."),
            "Eroare neașteptată.",
        )

    def test_localized_errors_are_preserved(self):
        self.assertEqual(
            error_message(UserFacingError("Format neacceptat"), "Eroare"),
            "Format neacceptat",
        )

    def test_unsupported_format(self):
        with self.assertRaisesRegex(UserFacingError, "Format neacceptat.*"):
            main.read_document("document.txt")

    def test_type_mismatch(self):
        message = main.validate_same_type("a.pdf", "b.docx")
        self.assertIn("Tipuri de fișiere diferite!", message)
        self.assertIn("Comparația este permisă", message)
        self.assertIsNone(main.validate_same_type("a.PDF", "b.pdf"))

    def test_worker_progress_and_success(self):
        worker = main.CompareWorker("a.docx", "b.docx")
        progress, results, errors = [], [], []
        worker.progress.connect(progress.append)
        worker.finished.connect(results.append)
        worker.error.connect(lambda *args: errors.append(args))
        with patch.object(main, "read_document", side_effect=[["Text"], ["Text"]]):
            worker.run()
        self.assertEqual(
            progress,
            [
                "Se citește documentul 1...",
                "Se citește documentul 2...",
                "Se compară documentele...",
            ],
        )
        self.assertEqual(results[0]["similarity"], 1.0)
        self.assertEqual(errors, [])

    def test_empty_documents(self):
        for paragraphs, number in [([[], ["Text"]], 1), ([["Text"], []], 2)]:
            worker = main.CompareWorker("a.pdf", "b.pdf")
            errors = []
            worker.error.connect(lambda *args: errors.append(args))
            with patch.object(main, "read_document", side_effect=paragraphs):
                worker.run()
            self.assertIn(f"documentul {number}", errors[0][0])
            self.assertIn("Verificați", errors[0][0])
            self.assertEqual(errors[0][1], "")

    def test_worker_keeps_technical_details_separate(self):
        worker = main.CompareWorker("a.pdf", "b.pdf")
        errors = []
        worker.error.connect(lambda *args: errors.append(args))
        with patch.object(
            main, "read_document", side_effect=RuntimeError("English diagnostic")
        ):
            worker.run()
        self.assertIn("Comparația nu a putut fi finalizată", errors[0][0])
        self.assertEqual(errors[0][1], "English diagnostic")

    def test_notification_buttons_and_detail_toggle(self):
        def inspect(box):
            self.assertEqual(box.text(), "Fișier salvat: <document.pdf>")
            self.assertEqual(box.textFormat(), Qt.PlainText)
            self.assertEqual(box.button(QMessageBox.Ok).text(), "Închide")
            self.assertEqual(box.detailedText(), "English diagnostic")
            button = next(
                b for b in box.buttons() if box.buttonRole(b) == QMessageBox.ActionRole
            )
            self.assertEqual(button.text(), "Afișează detaliile tehnice")
            button.click()
            self.assertEqual(button.text(), "Ascunde detaliile tehnice")
            button.click()
            self.assertEqual(button.text(), "Afișează detaliile tehnice")
            return QMessageBox.Ok

        with patch.object(QMessageBox, "exec_", inspect):
            show_notification(
                None,
                QMessageBox.Information,
                "Export reușit",
                "Fișier salvat: <document.pdf>",
                "English diagnostic",
            )

    def test_qt_overwrite_and_confirmation_buttons(self):
        self.assertEqual(
            QCoreApplication.translate(
                "QFileDialog", "%1 already exists.\nDo you want to replace it?"
            ),
            "%1 există deja.\nDoriți să îl înlocuiți?",
        )
        box = QMessageBox(
            QMessageBox.Question,
            "Confirmare",
            "Înlocuiți fișierul?",
            QMessageBox.Yes | QMessageBox.No,
        )
        self.assertEqual(box.button(QMessageBox.Yes).text(), "&Da")
        self.assertEqual(box.button(QMessageBox.No).text(), "&Nu")
        box.deleteLater()

    def test_file_dialog_is_localized(self):
        dialog = QFileDialog(None, "Salvează raportul")
        dialog.setAcceptMode(QFileDialog.AcceptSave)
        self.assertTrue(QApplication.testAttribute(Qt.AA_DontUseNativeDialogs))
        self.assertEqual(dialog.labelText(QFileDialog.Accept), "&Salvează")
        self.assertEqual(dialog.labelText(QFileDialog.Reject), "Anulează")
        dialog.deleteLater()

    def test_real_overwrite_notification_is_localized(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "existent.html"
            path.write_text("Raport existent")
            dialog = QFileDialog(None, "Salvează raportul", directory)
            dialog.setAcceptMode(QFileDialog.AcceptSave)
            dialog.selectFile(str(path))
            observed = []

            def dismiss_warning():
                for widget in QApplication.topLevelWidgets():
                    if isinstance(widget, QMessageBox) and widget.isVisible():
                        yes_button = widget.button(QMessageBox.Yes)
                        observed.append(
                            (widget.text(), yes_button.text() if yes_button else "")
                        )
                        no_button = widget.button(QMessageBox.No)
                        if no_button:
                            no_button.click()
                        else:
                            widget.reject()

            timer = QTimer()
            timer.timeout.connect(dismiss_warning)
            timer.start(10)
            timeout = QTimer(dialog)
            timeout.setSingleShot(True)
            timeout.timeout.connect(dialog.reject)
            timeout.start(2000)
            dialog.accept()
            timeout.stop()
            timer.stop()
            dialog.close()
            self.assertEqual(len(observed), 1)
            self.assertIn("există deja", observed[0][0])
            self.assertIn("Doriți să îl înlocuiți?", observed[0][0])
            self.assertEqual(observed[0][1], "&Da")
            self.assertEqual(path.read_text(), "Raport existent")

    def make_window(self):
        window = main.CompareWindow()
        window._result = main.compare_documents(["Un text."], ["Alt text."])
        window._result.update(path1="a.docx", path2="b.docx")
        self.addCleanup(window.close)
        return window

    def wait_export(self, window):
        started = time.monotonic()
        while window._busy and time.monotonic() - started < 10:
            QTest.qWait(20)
        self.assertFalse(window._busy)

    def test_inline_warning_and_loading_status(self):
        window = self.make_window()
        window.drop_left._set_file_display("a.pdf")
        self.assertEqual(window.drop_left.hint_lbl.text(), "Fișier încărcat")
        window.drop_right._set_file_display("b.docx")
        window._on_file_loaded("")
        self.assertIn("Comparația este permisă", window.status_lbl.text())
        self.assertFalse(window.compare_btn.isEnabled())

    def test_html_export_success(self):
        window = self.make_window()
        with tempfile.TemporaryDirectory() as directory:
            path = str(Path(directory) / "raport.html")
            with patch.object(
                QFileDialog, "getSaveFileName", return_value=(path, "")
            ), patch.object(main, "show_notification") as notify:
                window._export_report()
                self.wait_export(window)
            self.assertIn('<html lang="ro">', Path(path).read_text())
            self.assertEqual(notify.call_args.args[2], "Export HTML reușit")
            self.assertIn(path, notify.call_args.args[3])

    def test_export_errors_are_localized(self):
        window = self.make_window()
        for method, function in [
            (window._export_report, "export_html_report"),
            (window._export_pdf, "export_pdf_report"),
        ]:
            with self.subTest(export=function):
                with patch.object(
                    QFileDialog, "getSaveFileName", return_value=("raport", "")
                ), patch.object(
                    main, function, side_effect=PermissionError("Permission denied")
                ), patch.object(
                    main, "show_notification"
                ) as notify:
                    method()
                    self.wait_export(window)
                self.assertIn(
                    "Accesul la fișier a fost refuzat", notify.call_args.args[3]
                )
                self.assertEqual(notify.call_args.args[4], "Permission denied")

    def test_pdf_export_success(self):
        window = self.make_window()
        with patch.object(
            QFileDialog, "getSaveFileName", return_value=("raport.pdf", "")
        ), patch.object(main, "export_pdf_report") as export, patch.object(
            main, "show_notification"
        ) as notify:
            window._export_pdf()
            self.wait_export(window)
        Path("raport.pdf").unlink(missing_ok=True)
        export.assert_called_once()
        self.assertEqual(notify.call_args.args[2], "Export PDF reușit")

    def test_cancel_export_is_silent(self):
        window = self.make_window()
        with patch.object(
            QFileDialog, "getSaveFileName", return_value=("", "")
        ), patch.object(main, "show_notification") as notify:
            window._export_report()
            window._export_pdf()
        notify.assert_not_called()

    def test_real_document_reading_and_exports(self):
        with tempfile.TemporaryDirectory() as directory:
            docx_path = str(Path(directory) / "document.docx")
            doc = Document()
            doc.add_paragraph("Text în limba română.")
            doc.save(docx_path)
            paragraphs = main.read_document(docx_path)
            self.assertIn("Text în limba română.", paragraphs)
            result = main.compare_documents(paragraphs, paragraphs)
            pdf_path = str(Path(directory) / "raport.pdf")
            main.export_pdf_report(docx_path, docx_path, result, pdf_path)
            self.assertTrue(Path(pdf_path).read_bytes().startswith(b"%PDF"))
            self.assertTrue(main.read_document(pdf_path))


if __name__ == "__main__":
    unittest.main()
