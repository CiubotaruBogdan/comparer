# Comparer

Compară două documente DOCX sau două PDF-uri cu text selectabil, în coloane alăturate, cu evidențierea diferențelor la nivel de paragraf și cuvânt.

## Funcționalități

- Încărcare prin tragerea fișierului în fereastră sau selecție din dialog.
- Citire DOCX din corp, tabele, tabele imbricate, antete și subsoluri. Antetele/subsolurile partajate sunt deduplicate după identitatea părții documentului, nu după text.
- Identificarea paragrafelor identice, modificate, adăugate și șterse, cu evidențierea cuvintelor modificate.
- Înălțime comună pentru fiecare pereche de paragrafe, inclusiv după redimensionarea coloanelor, și scroll sincronizat în pixeli.
- Liste virtualizate, layout calculat în loturi și cache-uri limitate, fără precalcularea tuturor înălțimilor în fluxul normal de afișare.
- **Comparație nouă** și inversarea documentelor.
- **Doar diferențele**, cu 0–10 paragrafe de context înainte și după modificări.
- **Căutare** fără diferențiere între majuscule/minuscule, evidențiere și navigare între paragrafele găsite. Căutarea se aplică paragrafelor vizibile; dezactivați filtrul pentru a căuta în întregul rezultat.
- Opțiuni pentru ignorarea spațiilor suplimentare, a majusculelor și a antetelor/subsolurilor. Se aleg înainte de comparație; textele originale rămân vizibile în rezultat.
- Memorarea ultimului dosar, geometriei ferestrei, opțiunilor de comparație și filtrului prin `QSettings`.
- Export HTML/PDF în fundal, cu mesaje de progres și anulare cooperativă. Rapoartele includ opțiunile folosite și explicația scorului.
- Fonturile Unicode DejaVu Sans sunt incluse în aplicație, astfel încât exportul PDF nu depinde de fonturile instalate în sistem.
- Exportul PDF împarte sincronizat paragrafele lungi pe pagini și repetă antetul coloanelor.
- Notificări și controale în română, cu diacritice. Mesajele tehnice originale sunt disponibile separat, prin **Afișează detaliile tehnice**.

> Filtrarea și căutarea schimbă numai afișarea. Exportul păstrează rezultatul complet al comparației, cu opțiunile de normalizare aplicate la procesare.

## Procesare sigură

Selecția documentelor și comenzile care pot suprapune procesări sunt dezactivate pe durata comparației/exportului. Anularea și închiderea ferestrei solicită întreruperea și păstrează worker-ul până la finalizarea sa, fără oprire forțată a firului de execuție.

Anularea este cooperativă: dacă o bibliotecă este în curs de citire sau algoritmul de potrivire execută o operație indivizibilă, aplicația așteaptă încheierea acelei operații. Exportul folosește un fișier temporar în dosarul destinație și înlocuiește fișierul final numai după succes; erorile și anularea nu suprascriu un raport existent.

## Similaritate

Scorul este calculat pe secvențe de paragrafe, prin `SequenceMatcher`. Paragrafele considerate identice după normalizarea aleasă contribuie la potrivire; un paragraf modificat nu este considerat identic, chiar dacă diferă printr-un singur cuvânt. Nu este un scor semantic și nici un procent direct al cuvintelor identice.

## Legendă

| Culoare | Semnificație |
|---|---|
| Roșu | Paragraf prezent doar în documentul 1 (șters) |
| Verde | Paragraf prezent doar în documentul 2 (adăugat) |
| Galben | Paragraf modificat, prezent în ambele documente |

## Instalare și rulare

Versiunile dependențelor directe sunt fixate pentru a facilita reproducerea mediului verificat. Este recomandată o instalare curată într-un mediu virtual; pachetele vechi `fpdf` și `fpdf2` nu trebuie instalate simultan.

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# Linux/macOS: source .venv/bin/activate
python -m pip install -r requirements.txt
python main.py
```

Python 3.12 este versiunea verificată pentru această aplicație.

## Shortcut-uri

| Shortcut | Acțiune |
|---|---|
| `Ctrl+E` | Export HTML |
| `Ctrl+P` | Export PDF |
| `←` / `→` | Navigare între grupurile de diferențe |
| `Escape` | Comparație nouă; în timpul procesării solicită anulare și resetare |
| `Ctrl+F` | Activează câmpul de căutare |
| `F3` / `Shift+F3` | Următorul/anteriorul paragraf găsit |

## Denumirea rapoartelor

Documentele sursă trebuie denumite `AAAALLZZ_nivelclasificare_titlu.docx`
sau `.pdf`, cu o dată calendaristică validă și unul dintre nivelurile exacte:
`N` (neclasificat), `SSv` (secret de serviciu), `S` (secret), `SS` (strict secret).

La export, numele propus este
`AAAALLZZ_nivelclasificare_xxx-RaportComparație.html` sau `.pdf`.
Data este cea a exportului, `xxx` este întotdeauna text fix, iar nivelul este
maximul nivelurilor identificate în denumirile sursă: `N < SSv < S < SS`.
De exemplu, pentru `20261001_N_Analiză.docx` și `20261002_S_Analiză.docx`,
exportul din 19 octombrie 2026 propune `20261019_S_xxx-RaportComparație.pdf`.

Dacă cel puțin o denumire este neconformă, aplicația afișează un avertisment
și adaugă `-draft` înaintea extensiei. Dacă celălalt document are un nivel
valid, acesta este păstrat: de exemplu, `SS` și un nivel necunoscut produc
`20261019_SS_xxx-RaportComparație-draft.pdf`. Avertismentul identifică
documentul neconform și explică faptul că nivelul acestuia nu a putut fi determinat.

Doar dacă niciun nivel nu poate fi identificat se folosește implicit `N`:
`20261019_N_xxx-RaportComparație-draft.pdf`. Marcajul implicit NECLASIFICAT
nu confirmă clasificarea conținutului. În toate cazurile draft, stabilirea și
aplicarea nivelului corect revin creatorului raportului. Numele propus poate
fi editat în dialogul de salvare.

## Build executabil

Pe Windows:

```bat
build.bat
```

Rezultat: `dist/Comparer.exe`. Scriptul oprește build-ul dacă instalarea sau PyInstaller returnează o eroare. `Comparer.spec` include iconurile și fonturile/licența acestora.

Scriptul instalează dependențele aplicației și PyInstaller. Repository-ul conține codul final, resursele și configurația de build, fără teste sau pipeline-uri CI.

## Structură

- `document_core.py`: citirea și comparația documentelor, opțiuni și anulare.
- `report_export.py`: randare HTML și paginare PDF.
- `main.py`: fereastra, metadatele, modelele/delegatele și worker-ele.
- `notifications.py` / `romanian_qt.py`: notificări și localizarea dialogurilor Qt.
- `assets/fonts/`: fonturile necesare exportului PDF și licența lor.

## Limite

- Sunt acceptate DOCX cu DOCX și PDF cu PDF; nu este comparată formatarea documentelor Word ca într-un editor cu urmărirea modificărilor.
- PDF-urile scanate/fără text selectabil și cele protejate cu parolă primesc mesaje explicite. **OCR nu este inclus**; documentele scanate necesită preprocesare separată.
- Ordinea textului din PDF este estimată din coordonatele blocurilor. Documentele complexe pe mai multe coloane, tabelele PDF și layout-urile neobișnuite pot necesita verificare manuală.
- Paragrafele foarte repetitive pot crește durata algoritmului; anularea nu întrerupe forțat o operație internă a unei biblioteci.

Fonturi: [licența DejaVu Sans inclusă](assets/fonts/DEJAVU-LICENSE.txt).
