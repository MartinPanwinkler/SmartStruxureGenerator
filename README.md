# SmartStruxure Beschriftungsmaker (Dat Design is nich so doll)

Dat Projekt maakt formatteerte Excel-Beschriftungen ut `.xlsx`-, `.xlsm`- oder `.csv`-Rohdaten. Een Kästchen warrt dorbi as kumplett rechteckig Block behannelt. Zellweerte, Stil, Rahmen, Merges, Riegenhöhen un Spaltenbreden warrt ut een Masterdatei kopeert un denn mit Rohdaten füllt.

## Resultat vun de Referenzanalyse

De Masterdatei liggt hier:

```text
templates/Beschriftung SmartStruxure ERR.xlsx
```

De kumplette Analysebericht kann jümmers nee maakt warrn:

```bash
python inspect_template.py "templates/Beschriftung SmartStruxure ERR.xlsx" --json template_analysis.json
```

Dat Blatt `ASP-301-1` hett 33 visuelle Bereiche. De eegentlichen wedderkehrenden Modulbeschriftungen fangt af Rieg 27 an. De Referenz bruukt dree Geometrien:

- 16 Kanäle: `B27:M36`, twospaltig as 8 + 8 Kanäle
- 8 Kanäle: `B93:M98`, twospaltig as 4 + 4 Kanäle
- 12 Kanäle: `B178:M185`, twospaltig as 6 + 6 Kanäle

De Datenblätter `ASP-301-1_DP` un `ASP0303_DP (2)` hebbt de Datenpunktlisten, sünd aver sülvst keen Beschriftungsvorlagen. `template_config.json` bruukt dorüm blots bestätigte Blöck ut `ASP-301-1`. Jeder Kanal hett dor explizite Zielzellen in `data_slots`; dorüm blifft dat twospaltige Layout erhalten. Dat Analyseskript gifft zusätzlich Merges, Maße un Zellinhalten ut.

Dit SmartStruxure-Exportformat warrt automaatsch erkannt. Dorbi gellt de fasten Bedüden ut de Referenzblätter: B = Typ, C = Modul bzw. Datenpunkt, D = Beschrievung, E = Modulnummer, F = Kanal un G = Quelle. Modulriejen un Datenpunktriejen warrt automaatsch ünnerscheedt. Een manuelle Spaltenzutoordnung is dorför nich nödig; se blifft blots as Fallback för afwiekende normale Tabellen oder CSV-Dateien verföögbar.

### Dokumentkopp un manuelle Platzholler

Vör de I/O-Kästchen warrt de Originalbereich `A1:N25` kumplett kopeert. He hett de beiden Spannungsversorgungsbereiche un den Automation-Server-Block. Projektspezifische Inhalte warrt bewusst nich ut een ole Vorlage övernahmen, sonnern as direkt in Excel editeerbare Platzholler utggeven:

```text
{{DOKUMENTTITEL}}
{{NETZTEIL_MODUL_1}}
{{NETZTEIL_TEXT_1}}
{{NETZTEIL_MODUL_2}}
{{NETZTEIL_TEXT_2}}
{{AS_MODUL}}
{{ANLAGE}}
{{AUTOMATION_SERVER}}
{{IP_ADRESSE}}
{{SUBNETZMASKE}}
{{AS_VERSION}}
{{BENUTZER}}
{{PASSWORT}}
```

De Platzholler blievt normale Excel-Zelltexte un köönt na de Generierung manuell överschreven warrn. Jemehr Positschonen staht zentral ünner `document_header.cell_values` in `template_config.json`.

## Installation för Entwickler

Empfohlen is Python 3.11 oder 3.12. Ünner Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python main.py
```

Mit de Fish-Shell warrt dat dorvör vörsehn Aktivierungsskript bruukt:

```fish
source .venv/bin/activate.fish
python main.py
```

Ganz ahn Aktivierung geiht ok:

```bash
.venv/bin/python main.py
```

Ünner Windows (PowerShell):

```powershell
py -3.12 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```

För de GUI mutt Tkinter dor ween. Prüfen:

```bash
python -c "import tkinter; print(tkinter.TkVersion)"
```

Debian/Ubuntu:

```bash
sudo apt install python3-tk
```

Arch Linux/CachyOS:

```bash
sudo pacman -S tk
```

Denn de virtuelle Ümgebung gegevenenfalls nee maken. Op dat während de Entwicklung bruukte Linux-System fehl `libtk8.6.so`; Kernlogik un Tests loopt dorvun unafhängig.

### Ünner Linux ahn GUI testen

```bash
source .venv/bin/activate
python cli.py example_input.xlsx linux_test.xlsx \
  --template "templates/Beschriftung SmartStruxure ERR.xlsx"
```

Denn `linux_test.xlsx` mit LibreOffice Calc oder Microsoft Excel open maken. Een al vörhannen Utgavedatei warrt mit Afsicht nich överschreven. Vör een neen Test entweder een neen Naam bruken oder de ole Testdatei bewusst wegdoon.

## Bruuk

1. Rohdatei utwählen. Bi Excel-Dateien kann denn dat Tabellenblatt wählt warrn.
   Alternativ de `.xlsx`-, `.xlsm`- oder `.csv`-Datei op dat Drag-and-drop-Feld trecken.
2. Master-Vorlage utwählen.
3. Utgaveweg fastleggen. Een vörhannen Datei warrt ut Sicherheitsgrünnen nich överschreven.
4. De Vörschau un automaatsche Spaltenerkennung prüfen. Fehlende oder falsche Felder över **Spalten toordnen** korrigeren. De Toordnung kann för desülve Spaltenstruktur spiekert warrn.
5. **Beschriftung genereren** wählen.

De Logdatei liggt ünner `logs/app.log`. De Masterdatei warrt blots leest un nie spiekert.

In de kompileerte Windows-Version liggt de Logdatei duurhaft ünner:

```text
%LOCALAPPDATA%\SmartStruxureGenerator\logs\app.log
```

## Template-Konfiguration

All unsichere fachliche Positschonen staht in `template_config.json`:

```json
{
  "templates": {
    "rtd_di_16": {
      "sheet": "ASP-301-1_DP",
      "range": "A1:G18",
      "header_cells": {
        "module_type": "C1"
      },
      "data_start_row": 3,
      "max_channels": 16,
      "columns": {
        "channel": "B",
        "datapoint": "C",
        "description": "D",
        "source": "E"
      }
    }
  }
}
```

- `range` is de kumplett to kopeerende Masterblock.
- `header_cells` hett absolute Zelladressen binnen dissen Masterblock.
- `data_start_row` is de absolute eerste Datenrieg in’t Masterblatt.
- `max_channels` stüert de Opdeelung. 24 Datensätz bi 16 Kanälen gifft 16 + 8.
- `data_slots` ordent bi de analyseerten twospaltigen Vorlagen jeden Kanal expliziten Zielzellen to.
- `preserve_template_fields` lett fasten Inhalte as de druckten Kanalnummern `01` bis `16` unverännert ut de Mastervorlage stahn.
- `columns` steiht wieder för eenfache eenspaltige Vorlagen to Verfügung.
- `template_mapping` ordent Weerte ut `Modultyp` een Template to.
- `default_template` warrt bruukt, wenn keen speziell Mapping passt.
- `group_by` bestimmt, bi welk Feldwessel een nee logisch Modul entsteiht.
- `box_row_spacing` bestimmt de leddigen Riegen twüschen de kopeerten Blöck.

Vörhannen `{{ASP}}`-, `{{MODULE}}`- oder annere Platzholler in Textzellen warrt ersett. Alternativ funktioneert de expliziten Angaven in `header_cells` un `columns` ok ahn Platzholler in de Vorlage.

Een neen Kästchentyp warrt ergänzt, indem een wieder Inrag ünner `templates` anleggt un ünner `template_mapping` een Modultyp toordent warrt. Python-Code mutt dorför nich ännert warrn.

## CSV-Dateien

De Leser erkennt de Trenntekens Semikolon, Komma, Tabulator un Pipe. Bi de Tekencode warrt UTF-8 mit un ahn BOM, Windows-1252 un Latin-1 prüft. Excel-Weerte warrt as Weerte inleest; Formeln in Rohdaten warrt nich berekent.

## Bispeel un Tests

`example_input.xlsx` warrt mit folgen Kommando nee maakt:

```bash
python create_example_input.py
```

Tests utföhren:

```bash
pytest -q
```

Se prüft Spaltenerkennung, Gruppierung, 16-Kanal-Opdeelung, Blockkopie samt Stil/Merge/Maßen as ok den Schutz vör Överschrieven vun de Master- un Utgavedatei.

## Linux-Binary un Windows-EXE maken

Een eegenstännig Linux-Programm kann direkt ünner Linux boot warrn:

```bash
source .venv/bin/activate
pyinstaller --clean --noconfirm SmartStruxure_Beschriftungsgenerator.spec
```

Dat Linux-Programm liggt denn ünner `dist/SmartStruxure_Beschriftungsgenerator`. Tk/Tcl mutt vör den Build korrekt installiert ween, süss kann PyInstaller de GUI nich bündeln.

PyInstaller maakt Programme för dat Bedriefssystem, op dat dat löppt. De Linux-Build is dorüm **keen Windows-EXE**. Een native Windows-EXE warrt op Windows boot. Dat Repository dor utchecken un utföhren:

```powershell
py -3.12 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
pyinstaller --clean --noconfirm SmartStruxure_Beschriftungsgenerator.spec
```

De EXE liggt denn ünner:

```text
dist/SmartStruxure_Beschriftungsgenerator.exe
```

De Spec-Datei bündelt `template_config.json` un – sofern vörhannen – de Mastervorlage. För later licht editeerbare Konfigurationen kann `template_config.json` zusätzlich neven de EXE leggt warrn; disse externe Datei hett Vörrang. Datülve gellt för `templates/Beschriftung SmartStruxure ERR.xlsx` relativ to’n EXE-Ordner.

Alternativ maakt de enthollene GitHub-Actions-Workflow op een Windows-Runner automaatsch de `.exe`. Wine-basierte Cross-Builds sünd mööglich, aver düütlich fehleranfälliger un för dit Projekt nich empfohlen.

## GitHub-Repository anleggen

Na dat Prüfen vun lokale Dateien:

```bash
git init
git add .
git commit -m "Initial SmartStruxure generator"
git branch -M main
git remote add origin https://github.com/DEIN-NAME/SmartStruxureGenerator.git
git push -u origin main
```

De Referenzdatei kann interne Projektdaten hebben. Vör een öffentlichen Push schull prüft warrn, ob se veröffentlicht warrn dröff. Falls nich, `templates/*.xlsx` to de `.gitignore` todoon un de Vorlage trennt verdeelen.

## Projektstruktur

```text
main.py                         Programmeinstieg
cli.py                          Linux-/Kommandozeilentest ahn GUI
gui.py                          Tkinter-Böverflach, Vörschau un Mappingdialog
excel_reader.py                 CSV-/Excel-Inlesen un Sheet-Utwahl
mapping.py                      Alias-Erkennung, Normalisierung, Gruppierung
template_analyzer.py            Visuelle Analyse un Template-Utwahl
generator.py                    Blockkopie, Füllen un Utgave
models.py                       Datenmodelle
config.py                       Pade, Defaults un Logging
template_config.json            Korrigeerbare fachliche Zelltoordnung
inspect_template.py             Entwicklerwarktüüch
create_example_input.py         Maakt de Bispeel-Rohdatei
tests/                          Automaatsche Kernlogiktests
SmartStruxure_...spec           PyInstaller-Buildbeschrievung
```

## Bekannte Grenzen

- `openpyxl` bewahrt Zellformatierung, Merges, Maßen un normale Excel-Inhalte. Excel-Objekte as ActiveX-Steuerelemente, manche Teknungen oder externe Verbinnungen sünd nich Deel vun de kopeerte Kästchenlogik.
- `.xlsm` warrt as Rohdatenquelle ünnerstütt. De maakt Datei is mit Afsicht `.xlsx` un hett keen Makros.
