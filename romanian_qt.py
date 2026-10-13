"""Qt dialog translations bundled as Python so frozen builds need no .qm files."""

from PyQt5.QtCore import QLocale, QTranslator, Qt
from PyQt5.QtWidgets import QApplication


class RomanianQtTranslator(QTranslator):
    MESSAGES = {
        "QDialogButtonBox": {
            "OK": "Închide", "&OK": "Închide", "Save": "Salvează",
            "&Save": "&Salvează", "Open": "Deschide", "&Open": "&Deschide",
            "Cancel": "Anulează", "Close": "Închide", "&Close": "Închide",
            "Apply": "Aplică", "Reset": "Resetează", "Help": "Ajutor",
            "&Yes": "&Da", "Yes to &All": "Da pentru &toate",
            "&No": "&Nu", "N&o to All": "Nu pentru toate",
            "Abort": "Oprește", "Retry": "Reîncearcă", "Ignore": "Ignoră",
            "Restore Defaults": "Restabilește valorile implicite",
        },
        "QMessageBox": {
            "Show Details...": "Afișează detaliile tehnice",
            "Hide Details...": "Ascunde detaliile tehnice",
            "OK": "Închide", "Help": "Ajutor", "&Yes": "&Da", "&No": "&Nu",
        },
        "QFileDialog": {
            "Open": "Deschide", "Save As": "Salvează ca", "Find Directory": "Selectează dosarul",
            "Look in:": "Caută în:", "File &name:": "Nume &fișier:",
            "Files of type:": "Tip de fișier:", "Directory:": "Dosar:",
            "&Open": "&Deschide", "&Save": "&Salvează", "&Choose": "&Selectează",
            "Cancel": "Anulează", "All Files (*)": "Toate fișierele (*)",
            "All files (*)": "Toate fișierele (*)",
            "Back": "Înapoi", "Forward": "Înainte", "Parent Directory": "Dosarul părinte",
            "Create New Folder": "Creează un dosar", "List View": "Vizualizare listă",
            "Detail View": "Vizualizare detaliată", "Show ": "Afișează ",
            "Go back": "Înapoi", "Go forward": "Înainte",
            "Go to the parent directory": "Deschide dosarul părinte",
            "Create a New Folder": "Creează un dosar nou",
            "Change to list view mode": "Afișează ca listă",
            "Change to detail view mode": "Afișează detaliile",
            "Sidebar": "Bară laterală", "List of places and bookmarks": "Locații și marcaje",
            "Files": "Fișiere",
            "&Rename": "&Redenumește", "&Delete": "Șterge", "Show &hidden files": "Afișează fișierele ascunse",
            "&New Folder": "Dosar &nou", "New Folder": "Dosar nou",
            "File": "Fișier", "Directory": "Dosar", "Folder": "Dosar",
            "File not found": "Fișierul nu a fost găsit",
            "%1\nFile not found.\nPlease verify the correct file name was given.":
                "%1\nFișierul nu a fost găsit.\nVerificați dacă numele fișierului este corect.",
            "Directory not found": "Dosarul nu a fost găsit",
            "%1\nDirectory not found.\nPlease verify the correct directory name was given.":
                "%1\nDosarul nu a fost găsit.\nVerificați dacă numele dosarului este corect.",
            "%1 already exists.\nDo you want to replace it?":
                "%1 există deja.\nDoriți să îl înlocuiți?",
            "%1 is a directory.\nPlease choose a different file name.":
                "%1 este un dosar.\nAlegeți un alt nume pentru fișier.",
            "Delete": "Șterge", "Are you sure you want to delete '%1'?":
                "Sigur doriți să ștergeți «%1»?",
            "'%1' is write protected.\nDo you want to delete it anyway?":
                "«%1» este protejat la scriere.\nDoriți totuși să îl ștergeți?",
            "Could not delete directory.": "Dosarul nu a putut fi șters.",
            "The directory is not empty.": "Dosarul nu este gol.",
            "Could not rename file.": "Fișierul nu a putut fi redenumit.",
            "Could not create directory.": "Dosarul nu a putut fi creat.",
        },
        "QFileSystemModel": {
            "Name": "Nume", "Size": "Dimensiune", "Type": "Tip",
            "Date Modified": "Data modificării", "Folder": "Dosar",
            "File": "Fișier", "Computer": "Calculator", "Invalid filename": "Nume de fișier nevalid",
            "<b>The name \"%1\" can not be used.</b><p>Try using another name, with fewer characters or no punctuation marks.":
                "<b>Numele «%1» nu poate fi folosit.</b><p>Încercați un alt nume, cu mai puține caractere sau fără semne de punctuație.",
        },
    }

    def isEmpty(self):
        return False

    def translate(self, context, sourceText, disambiguation=None, n=-1):
        if context == "QPlatformTheme":
            context = "QDialogButtonBox"
        return self.MESSAGES.get(context, {}).get(sourceText, sourceText)


def install_romanian_dialogs(app):
    """Avoid OS-native dialogs that ignore application translations."""
    QApplication.setAttribute(Qt.AA_DontUseNativeDialogs, True)
    QLocale.setDefault(QLocale(QLocale.Romanian, QLocale.Romania))
    translator = RomanianQtTranslator(app)
    app.installTranslator(translator)
    # Retain the Python subclass for the entire application lifetime.
    app.romanian_translator = translator
