"""Romanian notifications, independent of the operating system language."""

from zipfile import BadZipFile

import fitz
from docx.opc.exceptions import PackageNotFoundError
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMessageBox


class UserFacingError(ValueError):
    """An application error whose message is already localized."""


def error_message(error: Exception, fallback: str) -> str:
    """Keep library/OS errors out of the user-facing notification text."""
    if isinstance(error, UserFacingError):
        return str(error)
    if isinstance(error, FileNotFoundError):
        return "Fișierul nu a fost găsit. Verificați dacă a fost mutat sau șters."
    if isinstance(error, PermissionError):
        return (
            "Accesul la fișier a fost refuzat. Verificați permisiunile și dacă "
            "fișierul este deschis în altă aplicație."
        )
    if isinstance(error, IsADirectoryError):
        return "Calea selectată este un dosar. Selectați un fișier."
    if isinstance(error, fitz.EmptyFileError):
        return "Fișierul PDF este gol. Selectați un document care conține text."
    if isinstance(error, (BadZipFile, PackageNotFoundError, fitz.FileDataError)):
        return (
            "Documentul nu poate fi citit. Verificați dacă fișierul există, "
            "nu este deteriorat și are formatul PDF sau DOCX corect."
        )
    if isinstance(error, OSError) and error.errno == 28:
        return "Nu există suficient spațiu pe disc. Eliberați spațiu și încercați din nou."
    return fallback


def show_notification(parent, icon, title: str, text: str, details: str = ""):
    """Show localized notification controls; retain technical details on demand."""
    box = QMessageBox(parent)
    box.setIcon(icon)
    box.setWindowTitle(title)
    box.setTextFormat(Qt.PlainText)
    box.setText(text)
    box.setStandardButtons(QMessageBox.Ok)
    box.button(QMessageBox.Ok).setText("Închide")
    if details:
        box.setDetailedText(details)
        # Qt creates an ActionRole button when detailed text is present.
        for button in box.buttons():
            if box.buttonRole(button) == QMessageBox.ActionRole:
                button.setText("Afișează detaliile tehnice")
                expanded = False

                def toggle_details(_checked=False, b=button):
                    nonlocal expanded
                    expanded = not expanded
                    b.setText("Ascunde detaliile tehnice" if expanded else "Afișează detaliile tehnice")

                button.clicked.connect(toggle_details)
    return box.exec_()
