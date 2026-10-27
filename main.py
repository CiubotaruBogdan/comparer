#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Document Comparator v1.3
Compară documente la nivel de paragraf (PDF cu PDF, DOCX cu DOCX),
cu highlighting al diferențelor, statistici detaliate,
export PDF/HTML, și vizualizare optimizată pentru documente mari.
"""

import sys
import os
import tempfile
import html as html_lib
from pathlib import Path
from typing import Optional
from datetime import datetime

from notifications import UserFacingError, error_message, show_notification
from romanian_qt import install_romanian_dialogs
from document_core import (
    _iter_block_texts,
    read_document,
    compare_documents,
    ComparisonCancelled,
)
from report_export import export_html_report, export_pdf_report, ExportCancelled
from report_naming import suggest_report_name

from docx import Document as DocxDocument
import fitz  # PyMuPDF

if sys.platform == "win32":
    import ctypes

    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Comparer.1.3")

from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSplitter,
    QFrame,
    QFileDialog,
    QStackedWidget,
    QMessageBox,
    QSizePolicy,
    QProgressBar,
    QShortcut,
    QStyleOptionViewItem,
    QStyledItemDelegate,
    QListView,
    QDialog,
    QGroupBox,
    QGridLayout,
    QAbstractItemView,
    QCheckBox,
    QLineEdit,
    QSpinBox,
)
from PyQt5.QtCore import (
    Qt,
    pyqtSignal,
    QThread,
    pyqtSlot,
    QTimer,
    QAbstractListModel,
    QModelIndex,
    QSize,
    QRect,
    QSettings,
)
from PyQt5.QtGui import (
    QFont,
    QColor,
    QPalette,
    QIcon,
    QKeySequence,
    QPainter,
    QTextDocument,
    QAbstractTextDocumentLayout,
)

# QPrinter no longer needed - PDF export uses fpdf2 directly


def _resource(relative: str) -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative
    return Path(__file__).parent / relative


def _load_icon() -> QIcon:
    for name in ("icon.png", "icon.ico"):
        p = _resource(name)
        if p.exists():
            return QIcon(str(p))
    return QIcon()


# ══════════════════════════════════════════════════════════════════════════════
# 1. Document Readers
# ══════════════════════════════════════════════════════════════════════════════


def validate_same_type(path1: str, path2: str) -> Optional[str]:
    """Validate that both files have the same extension. Returns error message or None."""
    ext1 = Path(path1).suffix.lower()
    ext2 = Path(path2).suffix.lower()
    if ext1 != ext2:
        return (
            f"Tipuri de fișiere diferite!\n\n"
            f"Documentul 1: {ext1.upper().lstrip('.')}\n"
            f"Documentul 2: {ext2.upper().lstrip('.')}\n\n"
            f"Comparația este permisă doar între fișiere de același tip:\n"
            f"  \u2022 PDF cu PDF\n"
            f"  \u2022 DOCX cu DOCX"
        )
    return None


# ══════════════════════════════════════════════════════════════════════════════
# 2. Motor de comparație
# ══════════════════════════════════════════════════════════════════════════════


import re as _re
import os as _os


def _strip_html_tags(text: str) -> str:
    """Remove HTML tags and decode common entities to get plain text."""
    clean = _re.sub(r"<[^>]+>", "", text)
    clean = (
        clean.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
        .replace("&nbsp;", " ")
        .replace("&#x27;", "'")
        .replace("&quot;", '"')
    )
    return clean.strip()


# ══════════════════════════════════════════════════════════════════════════════
# 2b. Extragere metadate documente
# ══════════════════════════════════════════════════════════════════════════════


def get_document_metadata(path: str) -> dict:
    """Extract metadata (author, creation date, modification date, etc.) from a document."""
    meta = {
        "file_name": Path(path).name,
        "file_path": str(path),
        "file_size": "",
        "format": Path(path).suffix.upper().lstrip("."),
        "author": "",
        "last_modified_by": "",
        "created": "",
        "modified": "",
        "revision": "",
        "title": "",
        "subject": "",
        "pages": "",
        "words": "",
    }

    # File size
    try:
        size_bytes = _os.path.getsize(path)
        if size_bytes < 1024:
            meta["file_size"] = f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            meta["file_size"] = f"{size_bytes / 1024:.1f} KB"
        else:
            meta["file_size"] = f"{size_bytes / (1024*1024):.2f} MB"
    except OSError:
        pass

    # File system modification time
    try:
        mtime = _os.path.getmtime(path)
        meta["file_modified"] = datetime.fromtimestamp(mtime).strftime(
            "%d.%m.%Y  %H:%M:%S"
        )
    except OSError:
        meta["file_modified"] = ""

    ext = Path(path).suffix.lower()

    if ext == ".docx":
        try:
            doc = DocxDocument(path)
            props = doc.core_properties
            meta["author"] = props.author or ""
            meta["last_modified_by"] = props.last_modified_by or ""
            meta["created"] = (
                props.created.strftime("%d.%m.%Y  %H:%M:%S") if props.created else ""
            )
            meta["modified"] = (
                props.modified.strftime("%d.%m.%Y  %H:%M:%S") if props.modified else ""
            )
            meta["revision"] = str(props.revision) if props.revision else ""
            meta["title"] = props.title or ""
            meta["subject"] = props.subject or ""
            # Word count from paragraphs
            body_texts = _iter_block_texts(doc.element.body, doc)
            word_count = sum(len(t.split()) for t in body_texts)
            meta["words"] = str(word_count)
            # Page count (from extended properties if available)
            try:
                from docx.opc.constants import RELATIONSHIP_TYPE as RT

                app_part = doc.part.package.part_related_by(RT.EXTENDED_PROPERTIES)
                if app_part is not None:
                    from lxml import etree

                    app_xml = etree.fromstring(app_part.blob)
                    ns = {
                        "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
                    }
                    pages_el = app_xml.find(".//ep:Pages", ns)
                    if pages_el is not None and pages_el.text:
                        meta["pages"] = pages_el.text
            except Exception:
                pass
        except Exception:
            pass

    elif ext == ".pdf":
        try:
            with fitz.open(path) as doc:
                meta["pages"] = str(doc.page_count)
                pdf_meta = doc.metadata
                if pdf_meta:
                    meta["author"] = pdf_meta.get("author", "") or ""
                    meta["title"] = pdf_meta.get("title", "") or ""
                    meta["subject"] = pdf_meta.get("subject", "") or ""
                    # PDF creation/modification dates
                    cr = pdf_meta.get("creationDate", "")
                    md = pdf_meta.get("modDate", "")
                    meta["created"] = _parse_pdf_date(cr) if cr else ""
                    meta["modified"] = _parse_pdf_date(md) if md else ""
                    meta["last_modified_by"] = pdf_meta.get("producer", "") or ""
                # Word count
                word_count = 0
                for page in doc:
                    text = page.get_text()
                    word_count += len(text.split())
                meta["words"] = str(word_count)
        except Exception:
            pass

    return meta


def _parse_pdf_date(date_str: str) -> str:
    """Parse PDF date format (D:YYYYMMDDHHmmSS) to readable format."""
    if not date_str:
        return ""
    # Remove D: prefix
    s = date_str
    if s.startswith("D:"):
        s = s[2:]
    # Remove timezone info
    for sep in ("+", "-", "Z"):
        if sep in s[8:]:
            s = s[: s.index(sep, 8)]
    try:
        if len(s) >= 14:
            dt = datetime.strptime(s[:14], "%Y%m%d%H%M%S")
            return dt.strftime("%d.%m.%Y  %H:%M:%S")
        elif len(s) >= 8:
            dt = datetime.strptime(s[:8], "%Y%m%d")
            return dt.strftime("%d.%m.%Y")
    except ValueError:
        pass
    return date_str


class DocumentInfoDialog(QDialog):
    """Dialog that shows metadata information about both compared documents."""

    def __init__(self, path1: str, path2: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Informații documente")
        self.setMinimumSize(700, 520)
        self.setStyleSheet("background:#fff;")
        self._build_ui(path1, path2)

    def _build_ui(self, path1: str, path2: str):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(16)

        # Title
        title = QLabel("Informații documente")
        f = QFont()
        f.setPointSize(16)
        f.setWeight(QFont.DemiBold)
        title.setFont(f)
        title.setStyleSheet("color:#111;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Get metadata for both documents
        meta1 = get_document_metadata(path1)
        meta2 = get_document_metadata(path2)

        # Document 1 group
        layout.addWidget(self._build_doc_group("Document 1 (Bază)", meta1))

        # Document 2 group
        layout.addWidget(self._build_doc_group("Document 2 (Comparație)", meta2))

        # Close button
        close_btn = QPushButton("Închide")
        close_btn.setFixedHeight(34)
        close_btn.setFixedWidth(120)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setStyleSheet("""
            QPushButton {
                background:#111; color:#fff;
                border:none; border-radius:4px; padding:6px 22px;
                font-size:10px;
            }
            QPushButton:hover { background:#333; }
            QPushButton:pressed { background:#000; }
        """)
        close_btn.clicked.connect(self.accept)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(close_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

    def _build_doc_group(self, title: str, meta: dict) -> QGroupBox:
        group = QGroupBox(title)
        group.setStyleSheet("""
            QGroupBox {
                font-size: 11px;
                font-weight: bold;
                color: #333;
                border: 1px solid #DDD;
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 14px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
                background: #fff;
            }
        """)

        grid = QGridLayout(group)
        grid.setContentsMargins(16, 12, 16, 12)
        grid.setSpacing(6)
        grid.setColumnMinimumWidth(0, 140)

        rows = [
            ("Nume fișier:", meta.get("file_name", "")),
            ("Format:", meta.get("format", "")),
            ("Dimensiune:", meta.get("file_size", "")),
            ("Pagini:", meta.get("pages", "")),
            ("Cuvinte:", meta.get("words", "")),
            ("Titlu document:", meta.get("title", "")),
            ("Subiect:", meta.get("subject", "")),
            ("Autor (creare):", meta.get("author", "")),
            ("Ultima editare de:", meta.get("last_modified_by", "")),
            ("Data creare:", meta.get("created", "")),
            ("Data ultimei editări:", meta.get("modified", "")),
            ("Revizie:", meta.get("revision", "")),
            ("Modificat pe disc:", meta.get("file_modified", "")),
        ]

        row_idx = 0
        for label_text, value in rows:
            if not value:
                continue
            lbl = QLabel(label_text)
            lbl.setStyleSheet("color:#888; font-size:10px; font-weight:normal;")
            val = QLabel(str(value))
            val.setStyleSheet("color:#111; font-size:10px; font-weight:normal;")
            val.setWordWrap(True)
            val.setTextInteractionFlags(Qt.TextSelectableByMouse)
            grid.addWidget(lbl, row_idx, 0, Qt.AlignTop)
            grid.addWidget(val, row_idx, 1, Qt.AlignTop)
            row_idx += 1

        if row_idx == 0:
            no_data = QLabel("Nu s-au putut extrage informații din acest document.")
            no_data.setStyleSheet(
                "color:#999; font-size:10px; font-style:italic; font-weight:normal;"
            )
            grid.addWidget(no_data, 0, 0, 1, 2)

        return group


# ══════════════════════════════════════════════════════════════════════════════
# 3. Thread de lucru
# ══════════════════════════════════════════════════════════════════════════════


class CompareWorker(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str, str)
    progress = pyqtSignal(str)
    settled = pyqtSignal()

    def __init__(self, path1, path2, options=None):
        super().__init__()
        self.path1, self.path2 = path1, path2
        self.options = options or {}

    def run(self):
        try:
            self.progress.emit("Se citește documentul 1...")
            paras1 = read_document(self.path1, cancel=self.isInterruptionRequested)
            if self.isInterruptionRequested():
                return
            self.progress.emit("Se citește documentul 2...")
            paras2 = read_document(self.path2, cancel=self.isInterruptionRequested)
            for number, paragraphs in ((1, paras1), (2, paras2)):
                if not paragraphs:
                    self.error.emit(
                        f"Nu s-au putut extrage paragrafe din documentul {number}.\n"
                        "Verificați dacă fișierul conține text și nu este scanat (imagine).",
                        "",
                    )
                    return
            if self.isInterruptionRequested():
                return
            self.progress.emit("Se compară documentele...")
            result = compare_documents(
                paras1,
                paras2,
                options=self.options,
                cancel=self.isInterruptionRequested,
            )
            if not self.isInterruptionRequested():
                result.update(
                    path1=self.path1, path2=self.path2, options=dict(self.options)
                )
                self.finished.emit(result)
        except ComparisonCancelled:
            pass
        except Exception as error:
            if not self.isInterruptionRequested():
                self.error.emit(
                    error_message(
                        error,
                        "Comparația nu a putut fi finalizată. Verificați documentele și încercați din nou.",
                    ),
                    str(error) if not isinstance(error, UserFacingError) else "",
                )
        finally:
            self.settled.emit()


class ExportWorker(QThread):
    completed = pyqtSignal(str, str)
    error = pyqtSignal(str, str, str)
    progress = pyqtSignal(str)

    def __init__(self, format_name, path, result):
        super().__init__()
        self.format_name, self.path, self.result = format_name, path, result

    def run(self):
        temporary = None
        try:
            if self.isInterruptionRequested():
                return
            self.progress.emit(f"Se exportă raportul {self.format_name}...")
            fd, temporary = tempfile.mkstemp(
                prefix=".comparer-", suffix=".tmp", dir=str(Path(self.path).parent)
            )
            os.close(fd)
            if self.format_name == "HTML":
                content = export_html_report(
                    self.result["path1"],
                    self.result["path2"],
                    self.result,
                    progress=self.progress.emit,
                    cancelled=self.isInterruptionRequested,
                )
                with open(temporary, "w", encoding="utf-8") as output:
                    output.write(content)
            else:
                export_pdf_report(
                    self.result["path1"],
                    self.result["path2"],
                    self.result,
                    temporary,
                    progress=self.progress.emit,
                    cancelled=self.isInterruptionRequested,
                )
            if self.isInterruptionRequested():
                return
            os.replace(temporary, self.path)
            temporary = None
            self.completed.emit(self.format_name, self.path)
        except ExportCancelled:
            pass
        except Exception as error:
            if not self.isInterruptionRequested():
                self.error.emit(
                    self.format_name,
                    error_message(
                        error,
                        f"Raportul {self.format_name} nu a putut fi salvat. Verificați locația aleasă și încercați din nou.",
                    ),
                    str(error),
                )
        finally:
            if temporary and os.path.exists(temporary):
                try:
                    os.unlink(temporary)
                except OSError:
                    pass


# ══════════════════════════════════════════════════════════════════════════════
# 4. DropZone widget
# ══════════════════════════════════════════════════════════════════════════════


class DropZone(QFrame):
    file_loaded = pyqtSignal(str)

    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.file_path: Optional[str] = None
        self._title = title
        self._last_dir = ""
        self.setAcceptDrops(True)
        self.setMinimumSize(280, 200)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self._setup_ui()
        self._apply_style(False)

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(6)
        layout.setContentsMargins(20, 16, 20, 16)

        self.title_lbl = QLabel(self._title)
        self.title_lbl.setAlignment(Qt.AlignCenter)
        f = QFont()
        f.setPointSize(10)
        f.setWeight(QFont.DemiBold)
        self.title_lbl.setFont(f)
        self.title_lbl.setStyleSheet("color:#111; background:transparent;")

        self.hint_lbl = QLabel("Trageți fișierul aici sau apăsați pentru a-l selecta")
        self.hint_lbl.setAlignment(Qt.AlignCenter)
        f2 = QFont()
        f2.setPointSize(8)
        self.hint_lbl.setFont(f2)
        self.hint_lbl.setStyleSheet("color:#BBB; background:transparent;")

        self.file_lbl = QLabel("")
        self.file_lbl.setAlignment(Qt.AlignCenter)
        f3 = QFont()
        f3.setPointSize(9)
        self.file_lbl.setFont(f3)
        self.file_lbl.setStyleSheet("color:#333; background:transparent;")
        self.file_lbl.setWordWrap(True)

        # Clear button (X) — hidden until a file is loaded
        self.clear_btn = QPushButton("\u2715")
        self.clear_btn.setFixedSize(22, 22)
        self.clear_btn.setCursor(Qt.PointingHandCursor)
        self.clear_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                color: #999;
                font-size: 14px;
                font-weight: bold;
                border-radius: 11px;
            }
            QPushButton:hover {
                background: #E0E0E0;
                color: #333;
            }
            QPushButton:pressed {
                background: #CCC;
            }
        """)
        self.clear_btn.hide()
        self.clear_btn.clicked.connect(self._clear_file)

        # Top row with clear button in top-right corner
        top_row = QHBoxLayout()
        top_row.setContentsMargins(0, 0, 0, 0)
        top_row.addStretch()
        top_row.addWidget(self.clear_btn)

        layout.addLayout(top_row)
        layout.addWidget(self.title_lbl)
        layout.addSpacing(2)
        layout.addWidget(self.hint_lbl)
        layout.addWidget(self.file_lbl)

    def _apply_style(self, loaded: bool):
        if loaded:
            self.setStyleSheet("""
                DropZone {
                    border: 1px solid #111;
                    border-radius: 6px;
                    background: #fff;
                }
                DropZone QLabel {
                    background: transparent;
                }
            """)
        else:
            self.setStyleSheet("""
                DropZone {
                    border: 1px dashed #BBBBBB;
                    border-radius: 6px;
                    background: #FAFAFA;
                }
                DropZone:hover {
                    border-color: #555;
                    background: #F4F4F4;
                }
                DropZone QLabel {
                    background: transparent;
                }
            """)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._browse()

    def _browse(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            f"Selectează {self._title}",
            self._last_dir,
            "Documente (*.docx *.pdf);;Word (*.docx);;PDF (*.pdf)",
        )
        if path:
            self._set_file(path)

    def _clear_file(self):
        """Clear the loaded file and reset the DropZone."""
        self.reset()
        self.file_loaded.emit("")  # Signal with empty path to update button state

    def _set_file_display(self, path: str):
        self.file_path = path
        self._last_dir = str(Path(path).parent)
        self.hint_lbl.setText("Fișier încărcat")
        self.hint_lbl.setStyleSheet("color:#AAA; background:transparent;")
        self.file_lbl.setText(Path(path).name)
        self.clear_btn.show()
        self._apply_style(True)

    def _set_file(self, path: str):
        self._set_file_display(path)
        self.file_loaded.emit(path)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            url = event.mimeData().urls()[0]
            if url.toLocalFile().lower().endswith((".docx", ".pdf")):
                event.acceptProposedAction()
                self.setStyleSheet("""
                    DropZone {
                        border: 1px solid #111;
                        border-radius: 6px;
                        background: #EFEFEF;
                    }
                    DropZone QLabel {
                        background: transparent;
                    }
                """)

    def dragLeaveEvent(self, event):
        self._apply_style(self.file_path is not None)

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            if path.lower().endswith((".docx", ".pdf")):
                self._set_file(path)

    def reset(self):
        self.file_path = None
        self.hint_lbl.setText("Trageți fișierul aici sau apăsați pentru a-l selecta")
        self.hint_lbl.setStyleSheet("color:#BBB; background:transparent;")
        self.file_lbl.setText("")
        self.clear_btn.hide()
        self._apply_style(False)


# ══════════════════════════════════════════════════════════════════════════════
# 5. Virtualized ComparePanel (memory-efficient for large documents)
# ══════════════════════════════════════════════════════════════════════════════


class BlockListModel(QAbstractListModel):
    """
    Model that holds block data (status, content, is_html, side) without
    creating any widgets. Only raw data is stored — rendering is done by
    the delegate on demand for visible items only.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._blocks: list = []  # list of (status, content, is_html)
        self._side: str = "left"
        self._selected_idx: int = -1

    def set_data(self, blocks: list, side: str):
        self.beginResetModel()
        self._blocks = blocks
        self._side = side
        self._selected_idx = -1
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return len(self._blocks)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        row = index.row()
        if row < 0 or row >= len(self._blocks):
            return None

        if role == Qt.UserRole:
            # Return full block tuple + side + selected
            status, content, is_html = self._blocks[row]
            return (status, content, is_html, self._side, row == self._selected_idx)
        return None

    def set_selected(self, idx: int):
        old = self._selected_idx
        self._selected_idx = idx
        if old >= 0 and old < len(self._blocks):
            i = self.index(old)
            self.dataChanged.emit(i, i, [Qt.UserRole])
        if idx >= 0 and idx < len(self._blocks):
            i = self.index(idx)
            self.dataChanged.emit(i, i, [Qt.UserRole])

    def block_count(self) -> int:
        return len(self._blocks)


class BlockDelegate(QStyledItemDelegate):
    """
    Custom delegate that paints paragraph blocks using QPainter directly.
    No widgets are created — this is the key to memory efficiency.
    Only visible items are rendered by the QListView viewport.
    """

    _BG = {
        "equal": QColor("#FFFFFF"),
        "replace": QColor("#FFECB3"),
        "delete": QColor("#FFCDD2"),
        "insert": QColor("#C8E6C9"),
        "placeholder": QColor("#F9F9F9"),
    }
    _BORDER_COLOR = {
        "equal": QColor("#E0E0E0"),
        "replace": QColor("#FFB300"),
        "delete": QColor("#E53935"),
        "insert": QColor("#43A047"),
        "placeholder": QColor("#EEEEEE"),
    }
    _SEL_BG = QColor("#EBEBEB")
    _SEL_BORDER = QColor("#888888")

    def __init__(self, parent=None):
        super().__init__(parent)
        self._font = QFont()
        self._font.setPointSize(10)
        self._doc_cache = {}  # LRU-like cache for QTextDocument (limited size)
        self._cache_max = 200  # max cached documents
        self._height_cache = {}
        self._search_text = ""

    def _get_doc(self, html_content: str, width: int) -> QTextDocument:
        """Get or create a QTextDocument for the given content and width."""
        key = (html_content, width)
        if key in self._doc_cache:
            return self._doc_cache[key]

        doc = QTextDocument()
        doc.setDefaultFont(self._font)
        doc.setTextWidth(width)
        doc.setHtml(html_content)

        # Evict oldest if cache full
        if len(self._doc_cache) >= self._cache_max:
            # Remove first item (oldest)
            first_key = next(iter(self._doc_cache))
            del self._doc_cache[first_key]

        self._doc_cache[key] = doc
        return doc

    def _build_html(self, status: str, content: str, is_html: bool, side: str) -> str:
        """Build the rich text HTML for a block."""
        if status == "placeholder":
            return ""

        badge = ""
        if status == "delete" and side == "left":
            badge = (
                '<span style="font-size:8px;background:#E53935;color:#fff;'
                'padding:1px 5px;border-radius:2px;margin-right:6px;">ȘTERS</span> '
            )
        elif status == "insert" and side == "right":
            badge = (
                '<span style="font-size:8px;background:#43A047;color:#fff;'
                'padding:1px 5px;border-radius:2px;margin-right:6px;">NOU</span> '
            )

        text = content if is_html else html_lib.escape(content)
        return badge + text

    def paint(
        self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex
    ):
        data = index.data(Qt.UserRole)
        if data is None:
            return

        status, content, is_html, side, selected = data
        rect = option.rect

        painter.save()
        painter.setClipRect(rect)

        # Background
        bg = self._SEL_BG if selected else self._BG.get(status, self._BG["equal"])
        painter.fillRect(rect, bg)

        # Left border (3px)
        border_color = (
            self._SEL_BORDER
            if selected
            else self._BORDER_COLOR.get(status, self._BORDER_COLOR["equal"])
        )
        border_rect = QRect(rect.left(), rect.top(), 3, rect.height())
        painter.fillRect(border_rect, border_color)

        # Content
        if status != "placeholder" and content:
            html_content = self._build_html(status, content, is_html, side)
            content_width = rect.width() - 20  # 10px padding on each side
            doc = self._get_doc(html_content, content_width)

            painter.translate(rect.left() + 13, rect.top() + 7)
            context = QAbstractTextDocumentLayout.PaintContext()
            selections = []
            if self._search_text:
                cursor = doc.find(self._search_text)
                while not cursor.isNull():
                    selection = QAbstractTextDocumentLayout.Selection()
                    selection.cursor = cursor
                    selection.format.setBackground(QColor("#FFF176"))
                    selection.format.setForeground(QColor("#111111"))
                    selections.append(selection)
                    cursor = doc.find(self._search_text, cursor)
            context.selections = selections
            doc.documentLayout().draw(painter, context)

        painter.restore()

    def _natural_height(self, block, side, width):
        status, content, is_html = block
        if status == "placeholder" or not content:
            return 32
        document = self._get_doc(
            self._build_html(status, content, is_html, side), max(20, width - 20)
        )
        return max(32, int(document.size().height()) + 14)

    def sizeHint(self, option: QStyleOptionViewItem, index: QModelIndex) -> QSize:
        panel = self.parent()
        row = index.row()
        if not index.isValid() or row >= panel._model.block_count():
            return QSize(100, 32)
        width = max(40, panel.viewport().width())
        partner = panel._partner
        partner_width = max(40, partner.viewport().width()) if partner else width
        key = (row, width, partner_width)
        if key not in self._height_cache:
            height = self._natural_height(
                panel._model._blocks[row], panel._model._side, width
            )
            if partner and row < partner._model.block_count():
                height = max(
                    height,
                    self._natural_height(
                        partner._model._blocks[row], partner._model._side, partner_width
                    ),
                )
            if len(self._height_cache) >= 2000:
                self._height_cache.clear()
            self._height_cache[key] = height
        return QSize(width, self._height_cache[key])

    def clear_cache(self):
        """Clear the document cache (call when loading new data)."""
        self._doc_cache.clear()
        self._height_cache.clear()


class ComparePanel(QListView):
    """
    Memory-efficient panel using QListView with a custom model and delegate.
    Only visible items are rendered — no widget is created per paragraph.
    This drastically reduces memory usage for large documents (thousands of paragraphs).
    """

    block_clicked = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setStyleSheet("""
            QListView {
                border: none;
                background: #FAFAFA;
            }
            QListView::item {
                margin: 1px 0px;
            }
        """)
        self.setSelectionMode(QListView.NoSelection)
        self.setEditTriggers(QListView.NoEditTriggers)
        self.setUniformItemSizes(False)
        self.setWordWrap(True)
        self.setSpacing(1)
        self.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.setLayoutMode(QListView.Batched)
        self.setBatchSize(100)

        self._model = BlockListModel(self)
        self._delegate = BlockDelegate(self)
        self.setModel(self._model)
        self.setItemDelegate(self._delegate)

        self.clicked.connect(self._on_item_clicked)

        self._partner: Optional["ComparePanel"] = None
        self._row_heights: list = []

    def load_blocks(self, blocks: list, side: str, diff_positions: list):
        self._delegate.clear_cache()
        self._model.set_data(blocks, side)
        self._row_heights = []

    def _compute_row_heights(self):
        """Compatibility helper for diagnostics; normal rendering measures lazily."""
        option = QStyleOptionViewItem()
        self._row_heights = [
            self._delegate.sizeHint(option, self._model.index(i)).height()
            for i in range(self._model.block_count())
        ]

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._delegate.clear_cache()
        if self._partner:
            self._partner._delegate.clear_cache()
            self._partner.scheduleDelayedItemsLayout()
        self.scheduleDelayedItemsLayout()

    def set_selected(self, idx: int):
        self._model.set_selected(idx)

    def scroll_to_block(self, idx: int):
        if 0 <= idx < self._model.block_count():
            index = self._model.index(idx)
            self.scrollTo(index, QListView.PositionAtTop)

    def get_block_count(self) -> int:
        return self._model.block_count()

    def _on_item_clicked(self, index: QModelIndex):
        if index.isValid():
            data = index.data(Qt.UserRole)
            if data and data[0] != "placeholder":
                self.block_clicked.emit(index.row())


# ══════════════════════════════════════════════════════════════════════════════
# 6. Fereastra principala
# ══════════════════════════════════════════════════════════════════════════════


def _flat_btn(text: str, dark: bool = False) -> QPushButton:
    btn = QPushButton(text)
    f = QFont()
    f.setPointSize(9)
    btn.setFont(f)
    btn.setCursor(Qt.PointingHandCursor)
    if dark:
        btn.setStyleSheet("""
            QPushButton {
                background:#111; color:#fff;
                border:none; border-radius:4px; padding:6px 22px;
            }
            QPushButton:hover   { background:#333; }
            QPushButton:pressed { background:#000; }
            QPushButton:disabled { background:#DDD; color:#AAA; }
        """)
    else:
        btn.setStyleSheet("""
            QPushButton {
                background:#fff; color:#111;
                border:1px solid #CCC; border-radius:4px; padding:5px 14px;
            }
            QPushButton:hover   { border-color:#555; }
            QPushButton:pressed { background:#EEE; }
            QPushButton:disabled { color:#CCC; border-color:#EEE; }
        """)
    return btn


class CompareWindow(QMainWindow):

    def __init__(self, settings=None):
        super().__init__()
        self.settings = (
            settings if settings is not None else QSettings("Comparer", "Comparer")
        )
        self._busy = False
        self._state = "idle"
        self._export_worker = None
        self._closing = False
        self._reset_pending = False
        self._visible_rows = []
        self._search_matches = []
        self._search_index = -1
        self.setWindowTitle("Comparer")
        self.setMinimumSize(1100, 680)
        self.setWindowIcon(_load_icon())
        self._syncing = False
        self._result: Optional[dict] = None
        self._worker: Optional[CompareWorker] = None
        self._current_diff = 0
        self._diff_count = 0
        self._sel_idx = -1
        self._setup_ui()
        self._setup_shortcuts()
        for zone in (self.drop_left, self.drop_right):
            zone._last_dir = self.settings.value("last_directory", "", type=str)
        geometry = self.settings.value("geometry")
        if geometry:
            self.restoreGeometry(geometry)
        self._set_busy(False)

    def _setup_shortcuts(self):
        QShortcut(QKeySequence("Ctrl+E"), self, self._export_report)
        QShortcut(QKeySequence("Ctrl+P"), self, self._export_pdf)
        QShortcut(QKeySequence("Right"), self, self._next_diff)
        QShortcut(QKeySequence("Left"), self, self._prev_diff)
        QShortcut(QKeySequence("Escape"), self, self._go_back)
        QShortcut(QKeySequence("Ctrl+F"), self, self._focus_search)
        QShortcut(QKeySequence("F3"), self, self._next_search)
        QShortcut(QKeySequence("Shift+F3"), self, self._prev_search)

    def _setup_ui(self):
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)
        self.stack.addWidget(self._build_welcome())
        self.stack.addWidget(self._build_compare())
        self.stack.setCurrentIndex(0)

    # ── Pagina 0: Welcome ─────────────────────────────────────────────────────

    def _build_welcome(self) -> QWidget:
        page = QWidget()
        page.setStyleSheet("background:#fff;")

        outer = QVBoxLayout(page)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addStretch(2)

        center = QWidget()
        center.setMaximumWidth(820)
        center.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        center_layout = QVBoxLayout(center)
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.setSpacing(0)

        center_layout.addSpacing(36)

        zones = QHBoxLayout()
        zones.setSpacing(16)
        self.drop_left = DropZone("Bază  —  Document 1")
        self.drop_right = DropZone("Comparație  —  Document 2")
        self.drop_left.setFixedHeight(148)
        self.drop_right.setFixedHeight(148)
        self.drop_left.file_loaded.connect(self._on_file_loaded)
        self.drop_right.file_loaded.connect(self._on_file_loaded)
        zones.addWidget(self.drop_left)
        zones.addWidget(self.drop_right)
        center_layout.addLayout(zones)
        center_layout.addSpacing(16)
        options_row = QHBoxLayout()
        self.ignore_whitespace = QCheckBox("Ignoră spațiile suplimentare")
        self.ignore_case = QCheckBox("Ignoră majusculele")
        self.ignore_headers = QCheckBox("Ignoră antetele și subsolurile")
        for key, checkbox in self._option_controls().items():
            checkbox.setChecked(self.settings.value(key, False, type=bool))
            checkbox.toggled.connect(self._save_preferences)
            options_row.addWidget(checkbox)
        center_layout.addLayout(options_row)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setFixedHeight(3)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar { background:#F0F0F0; border:none; border-radius:1px; }
            QProgressBar::chunk { background:#111; border-radius:1px; }
        """)
        self.progress_bar.hide()
        center_layout.addWidget(self.progress_bar)
        center_layout.addSpacing(8)

        self.status_lbl = QLabel("")
        self.status_lbl.setAlignment(Qt.AlignCenter)
        fst = QFont()
        fst.setPointSize(9)
        self.status_lbl.setFont(fst)
        self.status_lbl.setStyleSheet("color:#888;")
        self.status_lbl.setFixedHeight(18)

        self.compare_btn = _flat_btn("Compară", dark=True)
        fb = QFont()
        fb.setPointSize(10)
        self.compare_btn.setFont(fb)
        self.compare_btn.setFixedHeight(40)
        self.compare_btn.setMinimumWidth(200)
        self.compare_btn.setEnabled(False)
        self.compare_btn.clicked.connect(lambda: self._run_comparison())

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(self.compare_btn)
        self.cancel_compare_btn = _flat_btn("Anulează")
        self.cancel_compare_btn.clicked.connect(self._cancel_work)
        self.cancel_compare_btn.hide()
        btn_row.addWidget(self.cancel_compare_btn)
        btn_row.addStretch()

        center_layout.addWidget(self.status_lbl)
        center_layout.addSpacing(10)
        center_layout.addLayout(btn_row)

        h = QHBoxLayout()
        h.addStretch()
        h.addWidget(center)
        h.addStretch()
        outer.addLayout(h)
        outer.addStretch(3)

        return page

    # ── Pagina 1: Comparatie ──────────────────────────────────────────────────

    def _build_compare(self) -> QWidget:
        page = QWidget()
        page.setStyleSheet("background:#fff;")

        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        layout.addWidget(self._build_toolbar())
        layout.addWidget(self._build_filterbar())

        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.setStyleSheet(
            "QSplitter::handle { background:#DDDDDD; width:1px; }"
        )

        left_panel_w = QWidget()
        ll = QVBoxLayout(left_panel_w)
        ll.setContentsMargins(0, 0, 0, 0)
        ll.setSpacing(0)
        self.left_hdr = QLabel("  Document 1")
        self.left_hdr.setFixedHeight(28)
        self.left_hdr.setStyleSheet(
            "background:#111;color:#fff;font-size:10px;letter-spacing:.4px;"
            "font-family:system-ui,Arial,sans-serif;"
        )
        self.left_hdr.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        self.left_panel = ComparePanel()
        ll.addWidget(self.left_hdr)
        ll.addWidget(self.left_panel)

        right_panel_w = QWidget()
        rl = QVBoxLayout(right_panel_w)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(0)
        self.right_hdr = QLabel("  Document 2")
        self.right_hdr.setFixedHeight(28)
        self.right_hdr.setStyleSheet(
            "background:#444;color:#fff;font-size:10px;letter-spacing:.4px;"
            "font-family:system-ui,Arial,sans-serif;"
        )
        self.right_hdr.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        self.right_panel = ComparePanel()
        rl.addWidget(self.right_hdr)
        rl.addWidget(self.right_panel)

        self.splitter.addWidget(left_panel_w)
        self.splitter.addWidget(right_panel_w)
        self.splitter.setSizes([500, 500])
        self.left_panel._partner = self.right_panel
        self.right_panel._partner = self.left_panel
        layout.addWidget(self.splitter, 1)

        layout.addWidget(self._build_stats_panel())

        self.left_panel.block_clicked.connect(self._on_block_clicked)
        self.right_panel.block_clicked.connect(self._on_block_clicked)
        self.left_panel.verticalScrollBar().valueChanged.connect(self._sync_l2r)
        self.right_panel.verticalScrollBar().valueChanged.connect(self._sync_r2l)

        return page

    def _build_toolbar(self) -> QFrame:
        bar = QFrame()
        bar.setFixedHeight(46)
        bar.setStyleSheet("background:#fff; border-bottom:1px solid #E0E0E0;")

        lay = QHBoxLayout(bar)
        lay.setContentsMargins(12, 0, 12, 0)
        lay.setSpacing(8)

        self.new_btn = _flat_btn("Comparație nouă")
        self.new_btn.clicked.connect(self._go_back)
        self.swap_btn = _flat_btn("\u21c4  Inversează")
        self.export_btn = _flat_btn("Exportă HTML")
        self.export_pdf_btn = _flat_btn("Exportă PDF")
        self.info_btn = _flat_btn("\u2139  Informații documente")
        self.swap_btn.clicked.connect(self._swap_and_recompare)
        self.export_btn.clicked.connect(self._export_report)
        self.export_pdf_btn.clicked.connect(self._export_pdf)
        self.info_btn.clicked.connect(self._show_doc_info)

        self.prev_diff_btn = _flat_btn("\u2039 Diferență")
        self.next_diff_btn = _flat_btn("Diferență \u203a")
        self.diff_counter_lbl = QLabel("\u2014")
        fc = QFont()
        fc.setPointSize(9)
        self.diff_counter_lbl.setFont(fc)
        self.diff_counter_lbl.setStyleSheet("color:#666; min-width:52px;")
        self.diff_counter_lbl.setAlignment(Qt.AlignCenter)
        self.prev_diff_btn.clicked.connect(self._prev_diff)
        self.next_diff_btn.clicked.connect(self._next_diff)

        lay.addWidget(self.new_btn)
        lay.addWidget(self.swap_btn)
        lay.addSpacing(8)
        lay.addWidget(self.prev_diff_btn)
        lay.addWidget(self.diff_counter_lbl)
        lay.addWidget(self.next_diff_btn)
        lay.addStretch()
        lay.addWidget(self.export_btn)
        lay.addWidget(self.export_pdf_btn)
        lay.addSpacing(4)
        lay.addWidget(self.info_btn)

        return bar

    def _option_controls(self):
        return {
            "ignore_whitespace": self.ignore_whitespace,
            "ignore_case": self.ignore_case,
            "ignore_headers": self.ignore_headers,
        }

    def _comparison_options(self):
        return {
            key: checkbox.isChecked()
            for key, checkbox in self._option_controls().items()
        }

    def _save_preferences(self, *_args):
        for key, value in self._comparison_options().items():
            self.settings.setValue(key, value)
        if hasattr(self, "only_differences"):
            self.settings.setValue(
                "only_differences", self.only_differences.isChecked()
            )
            self.settings.setValue("context_rows", self.context_rows.value())

    def _build_filterbar(self):
        bar = QWidget()
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(12, 6, 12, 6)
        self.only_differences = QCheckBox("Doar diferențele")
        self.only_differences.setChecked(
            self.settings.value("only_differences", False, type=bool)
        )
        self.context_rows = QSpinBox()
        self.context_rows.setRange(0, 10)
        self.context_rows.setValue(self.settings.value("context_rows", 1, type=int))
        self.context_rows.setToolTip(
            "Paragrafe de context înainte și după fiecare diferență"
        )
        self.only_differences.setToolTip(
            "Filtrul afectează numai afișarea; exportul păstrează raportul complet."
        )
        self.only_differences.toggled.connect(self._apply_filter)
        self.context_rows.valueChanged.connect(self._apply_filter)
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Caută în documente… (Ctrl+F)")
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.textChanged.connect(self._update_search)
        self.search_edit.returnPressed.connect(self._next_search)
        self.search_previous = _flat_btn("‹")
        self.search_next = _flat_btn("›")
        self.search_previous.setToolTip("Rezultatul anterior (Shift+F3)")
        self.search_next.setToolTip("Rezultatul următor (F3)")
        self.search_previous.clicked.connect(self._prev_search)
        self.search_next.clicked.connect(self._next_search)
        self.search_count = QLabel("")
        self.cancel_export_btn = _flat_btn("Anulează exportul")
        self.cancel_export_btn.clicked.connect(self._cancel_work)
        self.cancel_export_btn.hide()
        self.export_status = QLabel("")
        layout.addWidget(self.only_differences)
        layout.addWidget(QLabel("Context:"))
        layout.addWidget(self.context_rows)
        layout.addWidget(self.search_edit, 1)
        layout.addWidget(self.search_previous)
        layout.addWidget(self.search_next)
        layout.addWidget(self.search_count)
        layout.addWidget(self.export_status)
        layout.addWidget(self.cancel_export_btn)
        return bar

    def _apply_filter(self, *_args):
        if not self._result:
            return
        count = len(self._result["left_blocks"])
        if self.only_differences.isChecked():
            keep = set()
            context = self.context_rows.value()
            for row, (left, right) in enumerate(
                zip(self._result["left_blocks"], self._result["right_blocks"])
            ):
                if left[0] not in ("equal", "placeholder") or right[0] not in (
                    "equal",
                    "placeholder",
                ):
                    keep.update(
                        range(max(0, row - context), min(count, row + context + 1))
                    )
            self._visible_rows = sorted(keep)
        else:
            self._visible_rows = list(range(count))
        self._original_to_visible = {
            original: visible for visible, original in enumerate(self._visible_rows)
        }
        self._sel_idx = -1
        self.left_panel.load_blocks(
            [self._result["left_blocks"][i] for i in self._visible_rows], "left", []
        )
        self.right_panel.load_blocks(
            [self._result["right_blocks"][i] for i in self._visible_rows], "right", []
        )
        self.export_status.setText(
            "Documentele sunt identice; nu există diferențe de afișat."
            if not self._visible_rows
            else ""
        )
        self.left_panel.scheduleDelayedItemsLayout()
        self.right_panel.scheduleDelayedItemsLayout()
        self._save_preferences()
        self._update_search()

    def _focus_search(self):
        if self.stack.currentIndex() == 1:
            self.search_edit.setFocus()
            self.search_edit.selectAll()

    def _update_search(self, *_args):
        query = self.search_edit.text().strip()
        self._search_matches = []
        if query and self._result:
            for row in self._visible_rows:
                texts = []
                for side in ("left_blocks", "right_blocks"):
                    _status, text, is_html = self._result[side][row]
                    texts.append(_strip_html_tags(text) if is_html else text)
                if any(query.casefold() in text.casefold() for text in texts):
                    self._search_matches.append(row)
        for panel in (self.left_panel, self.right_panel):
            panel._delegate._search_text = query
            panel.viewport().update()
        self._search_index = -1
        self.search_count.setText(
            f"{len(self._search_matches)} paragrafe" if query else ""
        )
        self.search_next.setEnabled(bool(self._search_matches))
        self.search_previous.setEnabled(bool(self._search_matches))
        if self._search_matches:
            self._next_search()

    def _move_search(self, direction):
        if not self._search_matches:
            return
        self._search_index = (self._search_index + direction) % len(
            self._search_matches
        )
        original = self._search_matches[self._search_index]
        row = self._original_to_visible[original]
        self._sel_idx = row
        self.left_panel.set_selected(row)
        self.right_panel.set_selected(row)
        self.left_panel.scroll_to_block(row)
        self.right_panel.scroll_to_block(row)
        self.search_count.setText(
            f"{self._search_index + 1}/{len(self._search_matches)} paragrafe"
        )

    def _next_search(self):
        self._move_search(1)

    def _prev_search(self):
        self._move_search(-1)

    def _build_stats_panel(self) -> QFrame:
        panel = QFrame()
        panel.setFixedHeight(76)
        panel.setStyleSheet("background:#F7F7F7; border-top:1px solid #E0E0E0;")

        lay = QHBoxLayout(panel)
        lay.setContentsMargins(20, 0, 20, 0)
        lay.setSpacing(0)

        sim_w = QWidget()
        sim_w.setStyleSheet("background:transparent;")
        sl = QVBoxLayout(sim_w)
        sl.setSpacing(0)
        sl.setContentsMargins(0, 0, 20, 0)
        self.sim_big = QLabel("\u2014")
        fb = QFont()
        fb.setPointSize(26)
        fb.setWeight(QFont.Light)
        self.sim_big.setFont(fb)
        self.sim_big.setStyleSheet("color:#111;")
        self.sim_big.setAlignment(Qt.AlignCenter)
        sl_lbl = QLabel("SIMILARITATE")
        score_help = "Scor calculat pe secvențe de paragrafe. Paragrafele modificate nu sunt considerate identice; opțiunile de normalizare se aplică înainte de comparație."
        sl_lbl.setToolTip(score_help)
        self.sim_big.setToolTip(score_help)
        fs = QFont()
        fs.setPointSize(7)
        sl_lbl.setFont(fs)
        sl_lbl.setStyleSheet("color:#AAA; letter-spacing:1px;")
        sl_lbl.setAlignment(Qt.AlignCenter)
        sl.addWidget(self.sim_big)
        sl.addWidget(sl_lbl)

        sep = QFrame()
        sep.setFrameShape(QFrame.VLine)
        sep.setStyleSheet("color:#DDDDDD;")
        sep.setFixedWidth(1)

        self.s_equal, w_eq = self._make_stat("IDENTICE", "\u2014")
        self.s_mod, w_mod = self._make_stat("MODIFICATE", "\u2014")
        self.s_added, w_add = self._make_stat("ADĂUGATE", "\u2014")
        self.s_deleted, w_del = self._make_stat("ȘTERSE", "\u2014")
        self.s_tot1, w_t1 = self._make_stat("PARAGRAFE DOC 1", "\u2014")
        self.s_tot2, w_t2 = self._make_stat("PARAGRAFE DOC 2", "\u2014")

        legend = self._build_legend()

        lay.addWidget(sim_w)
        lay.addWidget(sep)
        lay.addSpacing(4)
        for w in [w_eq, w_mod, w_add, w_del, w_t1, w_t2]:
            lay.addWidget(w)
        lay.addStretch()
        lay.addWidget(legend)

        return panel

    def _make_stat(self, label: str, value: str):
        w = QWidget()
        w.setStyleSheet("background:transparent;")
        lay = QVBoxLayout(w)
        lay.setSpacing(0)
        lay.setContentsMargins(14, 8, 14, 8)

        val = QLabel(value)
        fv = QFont()
        fv.setPointSize(15)
        fv.setWeight(QFont.DemiBold)
        val.setFont(fv)
        val.setStyleSheet("color:#111;")
        val.setAlignment(Qt.AlignCenter)

        lbl = QLabel(label)
        fl = QFont()
        fl.setPointSize(7)
        lbl.setFont(fl)
        lbl.setStyleSheet("color:#AAA; letter-spacing:.5px;")
        lbl.setAlignment(Qt.AlignCenter)

        lay.addWidget(val)
        lay.addWidget(lbl)
        return val, w

    def _build_legend(self) -> QWidget:
        w = QWidget()
        w.setStyleSheet("background:transparent;")
        lay = QVBoxLayout(w)
        lay.setSpacing(5)
        lay.setContentsMargins(8, 8, 0, 8)

        items = [
            ("#FFCDD2", "#E53935", "Șters  (doc. stânga)"),
            ("#C8E6C9", "#43A047", "Adăugat  (doc. dreapta)"),
            ("#FFECB3", "#FFB300", "Modificat  (ambele)"),
        ]
        fl = QFont()
        fl.setPointSize(8)
        for border, bg, txt in items:
            row = QHBoxLayout()
            row.setSpacing(7)
            swatch = QLabel()
            swatch.setFixedSize(22, 13)
            swatch.setStyleSheet(
                f"background:{bg};border-left:3px solid {border};border-radius:1px;"
            )
            lbl = QLabel(txt)
            lbl.setFont(fl)
            lbl.setStyleSheet("color:#888;")
            row.addWidget(swatch)
            row.addWidget(lbl)
            lay.addLayout(row)
        return w

    # ── Logica aplicatiei ────────────────────────────────────────────────────

    def _on_file_loaded(self, _path: str):
        if _path:
            self.settings.setValue("last_directory", str(Path(_path).parent))
        self._refresh_controls()

    def _refresh_controls(self):
        both = self.drop_left.file_path and self.drop_right.file_path
        mismatch = bool(
            both
            and validate_same_type(self.drop_left.file_path, self.drop_right.file_path)
        )
        self.compare_btn.setEnabled(not self._busy and bool(both) and not mismatch)
        for control in (
            self.drop_left,
            self.drop_right,
            *self._option_controls().values(),
        ):
            control.setEnabled(not self._busy)
        for control in (
            self.swap_btn,
            self.export_btn,
            self.export_pdf_btn,
            self.info_btn,
        ):
            control.setEnabled(not self._busy and self._result is not None)
        self.new_btn.setEnabled(not self._busy)
        self.only_differences.setEnabled(not self._busy)
        self.context_rows.setEnabled(not self._busy)
        if not self._busy:
            self.status_lbl.setText(
                "Tipuri diferite! Comparația este permisă doar PDF↔PDF sau DOCX↔DOCX."
                if mismatch
                else ""
            )
            self.status_lbl.setStyleSheet(
                "color:#E53935;" if mismatch else "color:#888;"
            )

    def _set_busy(self, busy):
        self._busy = busy
        if not busy:
            self._state = "idle"
        self.progress_bar.setVisible(busy and self._state != "exporting")
        self.cancel_compare_btn.setVisible(busy and self._state != "exporting")
        self.cancel_export_btn.setVisible(busy and self._state == "exporting")
        self._refresh_controls()

    def _run_comparison(self, path1=None, path2=None):
        if self._busy or (self._worker and self._worker.isRunning()):
            return
        p1 = path1 or self.drop_left.file_path
        p2 = path2 or self.drop_right.file_path
        if not p1 or not p2:
            return
        error = validate_same_type(p1, p2)
        if error:
            show_notification(
                self, QMessageBox.Warning, "Tipuri de fișiere incompatibile", error
            )
            return
        self._state = "running"
        self._set_busy(True)
        self.status_lbl.setText("Se procesează...")
        self._worker = CompareWorker(p1, p2, self._comparison_options())
        self._worker.finished.connect(self._on_comparison_done)
        self._worker.error.connect(self._on_comparison_error)
        self._worker.progress.connect(self.status_lbl.setText)
        self._worker.settled.connect(self._work_settled)
        self._worker.start()

    def _cancel_work(self):
        worker = self._export_worker or self._worker
        if worker and worker.isRunning():
            self._state = "cancelling"
            worker.requestInterruption()
            self.status_lbl.setText(
                "Se anulează; se așteaptă încheierea operației curente..."
            )
            self.export_status.setText("Se anulează...")
            self.cancel_compare_btn.setEnabled(False)
            self.cancel_export_btn.setEnabled(False)

    def _work_settled(self):
        workers = [worker for worker in (self._worker, self._export_worker) if worker]
        if any(worker.isRunning() for worker in workers):
            QTimer.singleShot(20, self._work_settled)
            return
        for worker in workers:
            worker.deleteLater()
        self._worker = None
        self._export_worker = None
        self._set_busy(False)
        self.cancel_compare_btn.setEnabled(True)
        self.cancel_export_btn.setEnabled(True)
        self.export_status.setText(
            "Documentele sunt identice; nu există diferențe de afișat."
            if self._result and not self._visible_rows
            else ""
        )
        if self._reset_pending:
            self._reset_pending = False
            self._go_back()
        if self._closing:
            self.close()

    @pyqtSlot(dict)
    def _on_comparison_done(self, result: dict):
        if self._closing or self._reset_pending or self._state == "cancelling":
            return
        self._result = result
        if not self._worker or not self._worker.isRunning():
            self._set_busy(False)
        self._update_compare_page(result)
        self.stack.setCurrentIndex(1)

    @pyqtSlot(str, str)
    def _on_comparison_error(self, msg: str, details: str = ""):
        if self._closing or self._state == "cancelling":
            return
        show_notification(
            self, QMessageBox.Critical, "Eroare la comparație", msg, details
        )

    def _on_block_clicked(self, idx: int):
        new_sel = -1 if self._sel_idx == idx else idx
        self._sel_idx = new_sel
        self.left_panel.set_selected(new_sel)
        self.right_panel.set_selected(new_sel)

    def _update_compare_page(self, result: dict):
        self._sel_idx = -1
        sim = result["similarity"]
        stats = result["stats"]
        sim_pct = f"{sim * 100:.1f}%"

        self.sim_big.setText(sim_pct)
        self.s_equal.setText(str(stats["equal"]))
        self.s_mod.setText(str(stats["modified"]))
        self.s_added.setText(str(stats["added"]))
        self.s_deleted.setText(str(stats["deleted"]))
        self.s_tot1.setText(str(stats["total_left"]))
        self.s_tot2.setText(str(stats["total_right"]))

        n1 = Path(result["path1"]).name
        n2 = Path(result["path2"]).name
        self.left_hdr.setText(f"  {n1}")
        self.right_hdr.setText(f"  {n2}")

        self._diff_count = result.get("diff_count", 0)
        self._current_diff = 0
        self._update_diff_nav()

        self._apply_filter()

    def _update_diff_nav(self):
        has = self._diff_count > 0
        self.prev_diff_btn.setEnabled(has and self._current_diff > 0)
        self.next_diff_btn.setEnabled(has and self._current_diff < self._diff_count - 1)
        self.diff_counter_lbl.setText(
            f"{self._current_diff + 1} / {self._diff_count}" if has else "0 / 0"
        )

    def _go_to_diff(self, idx: int):
        if not self._diff_count:
            return
        self._current_diff = max(0, min(idx, self._diff_count - 1))
        original = self._result["diff_positions"][self._current_diff]
        block_idx = self._original_to_visible.get(original)
        if block_idx is None:
            return
        self._syncing = True
        self.left_panel.scroll_to_block(block_idx)
        self.right_panel.scroll_to_block(block_idx)
        self._syncing = False
        self._update_diff_nav()

    def _next_diff(self):
        if self.stack.currentIndex() == 1:
            self._go_to_diff(self._current_diff + 1)

    def _prev_diff(self):
        if self.stack.currentIndex() == 1:
            self._go_to_diff(self._current_diff - 1)

    def _go_back(self):
        if self._busy:
            self._reset_pending = True
            self._cancel_work()
            return
        self._sel_idx = -1
        self._result = None
        self._diff_count = 0
        self._current_diff = 0
        self._visible_rows = []
        self._search_matches = []
        self.stack.setCurrentIndex(0)
        self.drop_left.reset()
        self.drop_right.reset()
        self.search_edit.clear()
        self.export_status.clear()
        self.left_panel.load_blocks([], "left", [])
        self.right_panel.load_blocks([], "right", [])
        self._update_diff_nav()
        self._set_busy(False)

    def closeEvent(self, event):
        self._save_preferences()
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.sync()
        if self._busy or any(
            worker and worker.isRunning()
            for worker in (self._worker, self._export_worker)
        ):
            self._closing = True
            self._cancel_work()
            event.ignore()
            return
        event.accept()

    def _swap_and_recompare(self):
        if not self._result or self._busy:
            return
        p1 = self._result["path1"]
        p2 = self._result["path2"]
        self.drop_left._set_file_display(p2)
        self.drop_right._set_file_display(p1)
        self._run_comparison(p2, p1)

    def _export_report(self):
        self._start_export("HTML")

    def _export_pdf(self):
        self._start_export("PDF")

    def _start_export(self, format_name):
        if not self._result or self._busy:
            return
        extension = format_name.lower()
        report_name = suggest_report_name(
            self._result["path1"], self._result["path2"], extension
        )
        if report_name.warning:
            show_notification(
                self,
                QMessageBox.Warning,
                "Clasificare nedeterminată — raport draft",
                report_name.warning,
            )
        path, _ = QFileDialog.getSaveFileName(
            self,
            f"Salvează raportul {format_name}",
            report_name.filename,
            f"{format_name} (*.{extension})",
        )
        if not path:
            return
        self._state = "exporting"
        self._set_busy(True)
        self._export_worker = ExportWorker(format_name, path, self._result)
        self._export_worker.progress.connect(self.export_status.setText)
        self._export_worker.completed.connect(self._on_export_done)
        self._export_worker.error.connect(self._on_export_error)
        self._export_worker.finished.connect(self._work_settled)
        self._export_worker.start()

    def _on_export_done(self, format_name, path):
        if not self._closing:
            show_notification(
                self,
                QMessageBox.Information,
                f"Export {format_name} reușit",
                f"Raportul {format_name} a fost salvat în:\n{path}",
            )

    def _on_export_error(self, format_name, text, details):
        if not self._closing:
            show_notification(
                self,
                QMessageBox.Critical,
                f"Eroare la exportul {format_name}",
                text,
                details,
            )

    def _show_doc_info(self):
        """Show document information dialog with metadata about both documents."""
        if not self._result:
            return
        dlg = DocumentInfoDialog(
            self._result["path1"],
            self._result["path2"],
            parent=self,
        )
        dlg.exec_()

    def _sync_l2r(self, value: int):
        if not self._syncing:
            self._syncing = True
            self.right_panel.verticalScrollBar().setValue(value)
            self._syncing = False

    def _sync_r2l(self, value: int):
        if not self._syncing:
            self._syncing = True
            self.left_panel.verticalScrollBar().setValue(value)
            self._syncing = False


# ══════════════════════════════════════════════════════════════════════════════
# 7. Entry point
# ══════════════════════════════════════════════════════════════════════════════


def main():
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    install_romanian_dialogs(app)
    app.setApplicationName("Comparer")
    app.setStyle("Fusion")
    app.setWindowIcon(_load_icon())

    palette = QPalette()
    palette.setColor(QPalette.Window, QColor("#FFFFFF"))
    palette.setColor(QPalette.WindowText, QColor("#111111"))
    palette.setColor(QPalette.Base, QColor("#FFFFFF"))
    palette.setColor(QPalette.AlternateBase, QColor("#F7F7F7"))
    palette.setColor(QPalette.Text, QColor("#111111"))
    palette.setColor(QPalette.ButtonText, QColor("#111111"))
    palette.setColor(QPalette.Button, QColor("#FFFFFF"))
    palette.setColor(QPalette.Highlight, QColor("#111111"))
    palette.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
    app.setPalette(palette)

    win = CompareWindow()
    if win.settings.value("geometry"):
        win.show()
    else:
        win.showMaximized()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
