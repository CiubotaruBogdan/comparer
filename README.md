# Comparer

Compară două documente DOCX sau PDF editabile side-by-side, cu evidențierea diferențelor la nivel de paragraf și cuvânt.

---

## Funcționalități

- **Drag & drop** sau click pentru încărcarea documentelor (`.docx`, `.pdf`)
- **Comparație la nivel de paragraf** — identifică paragrafele identice, modificate, adăugate și șterse
- **Comparație la nivel de cuvânt** — în paragrafele modificate, evidențiază exact cuvintele schimbate
- **Vizualizare side-by-side** cu scroll sincronizat
- **Scor de similaritate** — procentul de potrivire între cele două documente
- **Bară de statistici** — număr paragrafe identice / modificate / adăugate / șterse
- **Inversare documente** — re-rulează comparația cu documentele inversate
- **Export HTML** — salvează un raport HTML complet cu highlighting
- **Export PDF** — salvează un raport PDF cu titlu "Documente comparate: doc1, doc2" și data/ora raportului
- **Buton ștergere (✕)** — permite eliminarea unui document încărcat greșit fără a reseta toată interfața
- **Randare optimizată pentru memorie** — folosește QListView cu delegate custom (model/view pattern) în loc de widget-uri individuale per paragraf, reducând drastic consumul de memorie pe documente mari

## Legendă culori

| Culoare | Semnificație |
|---------|-------------|
| Roșu | Paragraf prezent doar în Documentul 1 (șters) |
| Verde | Paragraf prezent doar în Documentul 2 (adăugat) |
| Galben | Paragraf existent în ambele dar modificat |

---

## Cerințe

```
PyQt5 >= 5.15
python-docx >= 0.8.11
PyMuPDF >= 1.23
```

Instalare:

```bash
pip install -r requirements.txt
```

## Rulare

```bash
python main.py
```

## Shortcut-uri tastatură

| Shortcut | Acțiune |
|----------|--------|
| `Ctrl+E` | Export raport HTML |
| `Ctrl+P` | Export raport PDF |
| `←` / `→` | Navigare între diferențe |
| `Escape` | Comparatie nouă |

## Build executabil `.exe`

```bash
build.bat
```

Rezultat: `dist/Comparer.exe` — nu necesită Python sau dependențe pe mașina țintă.

> **Notă:** Comparația PDF funcționează doar pe PDF-uri text (editabile). PDF-urile scanate (imagine) necesită pre-procesare OCR.
