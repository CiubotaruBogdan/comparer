"""Classification and suggested report names derived from source filenames."""

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re


CLASSIFICATION_LEVELS = {"N": 0, "SSv": 1, "S": 2, "SS": 3}
CLASSIFICATION_NAMES = {
    "N": "NECLASIFICAT",
    "SSv": "SECRET DE SERVICIU",
    "S": "SECRET",
    "SS": "STRICT SECRET",
}
_DOCUMENT_NAME = re.compile(r"([0-9]{8})_(N|SSv|S|SS)_(.+)")


@dataclass(frozen=True)
class ReportName:
    filename: str
    classification: str
    invalid_sources: tuple[str, ...]

    @property
    def warning(self) -> str:
        if not self.invalid_sources:
            return ""
        files = "\n".join(f"• {name}" for name in self.invalid_sources)
        if len(self.invalid_sources) == 2:
            classification_notice = (
                "Nu s-a putut identifica nivelul de clasificare al niciunuia "
                "dintre documente. Noul fișier este denumit implicit "
                "NECLASIFICAT (N), cu sufixul «-draft». Această denumire nu "
                "confirmă că informațiile din raport sunt neclasificate. "
            )
        else:
            level = CLASSIFICATION_NAMES[self.classification]
            classification_notice = (
                "Nu se poate determina nivelul de clasificare al documentului "
                "enumerat mai sus. Nivelul cunoscut din celălalt document, "
                f"{level} ({self.classification}), este păstrat în denumirea "
                "raportului, cu sufixul «-draft». Nivelul final al raportului "
                "trebuie verificat, deoarece documentul cu denumire neconformă "
                "poate avea un nivel superior. "
            )
        return (
            "Următoarele documente nu au denumiri conforme cu formatul "
            "AAAALLZZ_nivelclasificare_titlu:\n\n"
            f"{files}\n\n"
            f"{classification_notice}"
            "Responsabilitatea stabilirii și aplicării nivelului corect de "
            "clasificare revine persoanei care creează raportul."
        )


def _parse_document_name(path: str) -> tuple[str, str] | None:
    match = _DOCUMENT_NAME.fullmatch(Path(path).stem)
    if match is None:
        return None
    stamp, classification, title = match.groups()
    try:
        datetime.strptime(stamp, "%Y%m%d")
    except ValueError:
        return None
    if not title.strip():
        return None
    return classification, title


def suggest_report_name(
    path1: str, path2: str, extension: str, *, today: date | None = None
) -> ReportName:
    """Keep the highest identified level and flag any unknown source as draft."""
    extension = extension.lower().lstrip(".")
    if extension not in ("html", "pdf"):
        raise ValueError("Formatul raportului trebuie să fie HTML sau PDF.")
    parsed = [_parse_document_name(path) for path in (path1, path2)]
    invalid = tuple(
        Path(path).name
        for path, document in zip((path1, path2), parsed)
        if document is None
    )
    classification = max(
        (document[0] for document in parsed if document is not None),
        key=CLASSIFICATION_LEVELS.get,
        default="N",
    )
    stamp = (today or date.today()).strftime("%Y%m%d")
    draft = "-draft" if invalid else ""
    return ReportName(
        f"{stamp}_{classification}_xxx-RaportComparație{draft}.{extension}",
        classification,
        invalid,
    )
