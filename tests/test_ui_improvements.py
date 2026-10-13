import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch
from PyQt5.QtCore import QSettings
from PyQt5.QtWidgets import QApplication, QShortcut
from PyQt5.QtTest import QTest
import main
from romanian_qt import install_romanian_dialogs


class UIImprovementsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        install_romanian_dialogs(cls.app)

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.settings = QSettings(
            str(Path(self.directory.name) / "prefs.ini"), QSettings.IniFormat
        )
        self.window = main.CompareWindow(settings=self.settings)
        self.window.show()
        QTest.qWait(60)

    def tearDown(self):
        for worker in (self.window._worker, self.window._export_worker):
            if worker:
                worker.requestInterruption()
                worker.wait(10000)
        self.window._work_settled()
        self.window.close()
        self.app.processEvents()
        self.directory.cleanup()

    def wait_until(self, predicate, timeout=10000):
        started = time.monotonic()
        while not predicate() and (time.monotonic() - started) * 1000 < timeout:
            QTest.qWait(20)
        self.assertTrue(predicate())

    def set_result(self, left, right):
        result = main.compare_documents(left, right)
        result.update(path1="a.docx", path2="b.docx")
        self.window._on_comparison_done(result)
        QTest.qWait(100)
        return result

    def test_new_comparison_and_escape(self):
        self.set_result(["A"], ["B"])
        self.assertIn(
            "Esc", [key.key().toString() for key in self.window.findChildren(QShortcut)]
        )
        self.window.new_btn.click()
        self.assertEqual(self.window.stack.currentIndex(), 0)
        self.assertIsNone(self.window._result)
        self.assertEqual(self.window.left_panel.get_block_count(), 0)

    def test_busy_stays_busy_on_file_event_and_duplicate_start(self):
        self.window.drop_left._set_file_display("a.docx")
        self.window.drop_right._set_file_display("b.docx")
        self.window._state = "running"
        self.window._set_busy(True)
        self.window._on_file_loaded("b.docx")
        self.assertFalse(self.window.compare_btn.isEnabled())
        self.assertFalse(self.window.drop_left.isEnabled())
        with patch.object(main, "CompareWorker") as worker:
            self.window._run_comparison()
        worker.assert_not_called()
        self.window._set_busy(False)

    def test_shared_heights_and_resize(self):
        self.set_result(
            ["Scurt.", "Al doilea.", "Final."],
            ["Paragraf foarte lung. " * 100, "Al doilea.", "Final."],
        )
        for size in [(1100, 680), (1300, 800)]:
            self.window.resize(*size)
            self.window.splitter.setSizes([350, 750])
            QTest.qWait(120)
            left, right = self.window.left_panel, self.window.right_panel
            left._compute_row_heights()
            right._compute_row_heights()
            self.assertEqual(left._row_heights, right._row_heights)
            self.assertEqual(
                left.visualRect(left._model.index(1)).top(),
                right.visualRect(right._model.index(1)).top(),
            )
            left.verticalScrollBar().setValue(120)
            self.assertEqual(
                left.verticalScrollBar().value(), right.verticalScrollBar().value()
            )

    def test_filter_context_and_original_navigation(self):
        self.set_result(["A", "B", "C", "D", "E"], ["A", "B", "NOU", "D", "E"])
        self.window.context_rows.setValue(0)
        self.window.only_differences.setChecked(True)
        self.assertEqual(self.window._visible_rows, [2])
        self.window._go_to_diff(0)
        self.assertEqual(self.window.left_panel.get_block_count(), 1)
        self.window.context_rows.setValue(1)
        self.assertEqual(self.window._visible_rows, [1, 2, 3])
        self.window.only_differences.setChecked(False)
        self.assertEqual(self.window.left_panel.get_block_count(), 5)

    def test_no_differences_filter(self):
        self.set_result(["Identic"], ["Identic"])
        self.window.only_differences.setChecked(True)
        self.assertEqual(self.window.left_panel.get_block_count(), 0)
        self.assertFalse(self.window.next_diff_btn.isEnabled())

    def test_search_original_text_case_insensitive_and_navigation(self):
        self.set_result(
            ["Un Text căutat.", "Alt text."], ["Un TEXT diferit.", "Alt text."]
        )
        self.window.search_edit.setText("text")
        self.assertEqual(self.window._search_matches, [0, 1])
        self.assertEqual(self.window._search_index, 0)
        self.window._next_search()
        self.assertEqual(self.window._search_index, 1)
        self.window._prev_search()
        self.assertEqual(self.window._search_index, 0)
        self.window.search_edit.clear()
        self.assertEqual(self.window._search_matches, [])
        self.assertEqual(self.window.left_panel._delegate._search_text, "")

    def test_preferences_restore(self):
        self.window.ignore_case.setChecked(True)
        self.window.ignore_whitespace.setChecked(True)
        self.window.settings.setValue("last_directory", self.directory.name)
        self.window._save_preferences()
        other = main.CompareWindow(settings=self.settings)
        try:
            self.assertTrue(other.ignore_case.isChecked())
            self.assertTrue(other.ignore_whitespace.isChecked())
            self.assertEqual(other.drop_left._last_dir, self.directory.name)
        finally:
            other.close()

    def test_export_worker_atomic_preserves_destination_on_error(self):
        result = self.set_result(["A"], ["B"])
        path = Path(self.directory.name) / "report.html"
        path.write_text("Original")
        worker = main.ExportWorker("HTML", str(path), result)
        errors = []
        worker.error.connect(lambda *args: errors.append(args))
        with patch.object(
            main, "export_html_report", side_effect=RuntimeError("failure")
        ):
            worker.run()
        self.assertEqual(path.read_text(), "Original")
        self.assertEqual(len(errors), 1)
        self.assertEqual(list(Path(self.directory.name).glob(".comparer-*")), [])

    def test_async_export_and_busy_cleanup(self):
        self.set_result(["A"], ["B"])
        path = str(Path(self.directory.name) / "report.html")
        with patch.object(
            main.QFileDialog, "getSaveFileName", return_value=(path, "")
        ), patch.object(main, "show_notification") as notify:
            self.window._export_report()
            self.assertTrue(self.window._busy)
            self.wait_until(lambda: not self.window._busy)
        self.assertTrue(Path(path).read_text().startswith("<!DOCTYPE html>"))
        self.assertEqual(notify.call_args.args[2], "Export HTML reușit")
        self.assertIsNone(self.window._export_worker)

    def test_cancel_compare_and_reset_without_destroying_running_thread(self):
        self.window.drop_left._set_file_display("a.docx")
        self.window.drop_right._set_file_display("b.docx")

        def slow_read(*args, **kwargs):
            for _ in range(30):
                if kwargs["cancel"]():
                    raise main.ComparisonCancelled()
                time.sleep(0.005)
            return ["Text"]

        with patch.object(main, "read_document", side_effect=slow_read):
            self.window._run_comparison()
            worker = self.window._worker
            self.window._go_back()
            self.assertIs(self.window._worker, worker)
            self.wait_until(lambda: not self.window._busy)
        self.assertIsNone(self.window._worker)
        self.assertIsNone(self.window._result)
        self.assertEqual(self.window.stack.currentIndex(), 0)

    def test_close_during_work_waits_for_thread(self):
        self.window.drop_left._set_file_display("a.docx")
        self.window.drop_right._set_file_display("b.docx")

        def slow_read(*args, **kwargs):
            time.sleep(0.08)
            return ["Text"]

        with patch.object(main, "read_document", side_effect=slow_read):
            self.window._run_comparison()
            self.window.close()
            self.assertTrue(self.window._closing)
            self.wait_until(lambda: not self.window.isVisible())
        self.assertIsNone(self.window._worker)


if __name__ == "__main__":
    unittest.main()
