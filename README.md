# SmartStruxure Beschriftungsgenerator

Das Projekt erzeugt formatierte Excel-Beschriftungen aus `.xlsx`-, `.xlsm`- oder `.csv`-Rohdaten. Ein Kästchen wird dabei als kompletter rechteckiger Block behandelt. Zellwerte, Styles, Rahmen, Merges, Zeilenhöhen und Spaltenbreiten werden aus einer Masterdatei kopiert und anschließend mit Rohdaten befüllt.

## Ergebnis der Referenzanalyse

Die Masterdatei liegt unter:

```text
templates/Beschriftung SmartStruxure ERR.xlsx
```

Der vollständige Analysebericht lässt sich jederzeit neu erzeugen:

```bash
python inspect_template.py "templates/Beschriftung SmartStruxure ERR.xlsx" --json template_analysis.json
```

Das Blatt `ASP-301-1` enthält 33 visuelle Bereiche. Die eigentlichen wiederkehrenden Modulbeschriftungen beginnen ab Zeile 27. Die Referenz verwendet drei Geometrien:

- 16 Kanäle: `B27:M36`, zweispaltig als 8 + 8 Kanäle
- 8 Kanäle: `B93:M98`, zweispaltig als 4 + 4 Kanäle
- 12 Kanäle: `B178:M185`, zweispaltig als 6 + 6 Kanäle

Die Datenblätter `ASP-301-1_DP` und `ASP0303_DP (2)` enthalten die Datenpunktlisten, sind aber selbst keine Beschriftungsvorlagen. `template_config.json` verwendet deshalb ausschließlich bestätigte Blöcke aus `ASP-301-1`. Jeder Kanal besitzt dort explizite Zielzellen in `data_slots`; dadurch bleibt das zweispaltige Layout erhalten. Das Analyseskript gibt zusätzlich Merges, Maße und Zellinhalte aus.

Dieses SmartStruxure-Exportformat wird automatisch erkannt. Dabei gelten die festen Bedeutungen aus den Referenzblättern: B = Typ, C = Modul beziehungsweise Datenpunkt, D = Beschreibung, E = Modulnummer, F = Kanal und G = Quelle. Modulzeilen und Datenpunktzeilen werden automatisch unterschieden. Eine manuelle Spaltenzuordnung ist dafür nicht erforderlich; sie bleibt nur als Fallback für abweichende normale Tabellen oder CSV-Dateien verfügbar.

### Dokumentkopf und manuelle Platzhalter

Vor den I/O-Kästchen wird der Originalbereich `A1:N25` vollständig kopiert. Er enthält die beiden Spannungsversorgungsbereiche und den Automation-Server-Block. Projektspezifische Inhalte werden bewusst nicht aus einer alten Vorlage übernommen, sondern als direkt in Excel editierbare Platzhalter ausgegeben:

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

Die Platzhalter bleiben normale Excel-Zelltexte und können nach der Generierung manuell überschrieben werden. Ihre Positionen stehen zentral unter `document_header.cell_values` in `template_config.json`.

## Installation für Entwickler

Empfohlen ist Python 3.11 oder 3.12. Unter Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python main.py
```

Mit der Fish-Shell wird das dafür vorgesehene Aktivierungsskript verwendet:

```fish
source .venv/bin/activate.fish
python main.py
```

Ganz ohne Aktivierung funktioniert auch:

```bash
.venv/bin/python main.py
```

Unter Windows (PowerShell):

```powershell
py -3.12 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```

Für die GUI muss Tkinter vorhanden sein. Prüfen:

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

Danach die virtuelle Umgebung gegebenenfalls neu erstellen. Auf dem während der Entwicklung verwendeten Linux-System fehlte `libtk8.6.so`; Kernlogik und Tests funktionieren davon unabhängig.

### Unter Linux ohne GUI testen

```bash
source .venv/bin/activate
python cli.py example_input.xlsx linux_test.xlsx \
  --template "templates/Beschriftung SmartStruxure ERR.xlsx"
```

Anschließend `linux_test.xlsx` mit LibreOffice Calc oder Microsoft Excel öffnen. Eine bereits vorhandene Ausgabedatei wird absichtlich nicht überschrieben. Vor einem erneuten Test entweder einen neuen Namen verwenden oder die alte Testdatei bewusst entfernen.

## Benutzung

1. Rohdatei auswählen. Bei Excel-Dateien kann anschließend das Tabellenblatt gewählt werden.
   Alternativ die `.xlsx`-, `.xlsm`- oder `.csv`-Datei auf das Drag-and-drop-Feld ziehen.
2. Master-Vorlage auswählen.
3. Ausgabepfad festlegen. Eine vorhandene Datei wird aus Sicherheitsgründen nicht überschrieben.
4. Die Vorschau und automatische Spaltenerkennung prüfen. Fehlende oder falsche Felder über **Spalten zuordnen** korrigieren. Die Zuordnung kann für dieselbe Spaltenstruktur gespeichert werden.
5. **Beschriftung generieren** wählen.

Die Logdatei liegt unter `logs/app.log`. Die Masterdatei wird nur gelesen und nie gespeichert.

In der kompilierten Windows-Version liegt die Logdatei dauerhaft unter:

```text
%LOCALAPPDATA%\SmartStruxureGenerator\logs\app.log
```

## Template-Konfiguration

Alle unsicheren fachlichen Positionen stehen in `template_config.json`:

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

- `range` ist der vollständig zu kopierende Masterblock.
- `header_cells` enthält absolute Zelladressen innerhalb dieses Masterblocks.
- `data_start_row` ist die absolute erste Datenzeile im Masterblatt.
- `max_channels` steuert die Aufteilung. 24 Datensätze bei 16 Kanälen ergeben 16 + 8.
- `data_slots` ordnet bei den analysierten zweispaltigen Vorlagen jeden Kanal expliziten Zielzellen zu.
- `preserve_template_fields` lässt feste Inhalte wie die gedruckten Kanalnummern `01` bis `16` unverändert aus der Mastervorlage stehen.
- `columns` steht weiterhin für einfache einspaltige Vorlagen zur Verfügung.
- `template_mapping` ordnet Werte aus `Modultyp` einem Template zu.
- `default_template` wird verwendet, wenn kein spezielles Mapping passt.
- `group_by` bestimmt, bei welchem Feldwechsel ein neues logisches Modul entsteht.
- `box_row_spacing` bestimmt die Leerzeilen zwischen den kopierten Blöcken.

Vorhandene `{{ASP}}`-, `{{MODULE}}`- oder andere Platzhalter in Textzellen werden ersetzt. Alternativ funktionieren die expliziten Angaben in `header_cells` und `columns` auch ohne Platzhalter in der Vorlage.

Ein neuer Kästchentyp wird ergänzt, indem ein weiterer Eintrag unter `templates` angelegt und unter `template_mapping` einem Modultyp zugeordnet wird. Python-Code muss dafür nicht geändert werden.

## CSV-Dateien

Der Leser erkennt die Trennzeichen Semikolon, Komma, Tabulator und Pipe. Bei der Zeichenkodierung werden UTF-8 mit und ohne BOM, Windows-1252 und Latin-1 geprüft. Excel-Werte werden als Werte eingelesen; Formeln in Rohdaten werden nicht berechnet.

## Beispiel und Tests

`example_input.xlsx` wird mit folgendem Befehl neu erstellt:

```bash
python create_example_input.py
```

Tests ausführen:

```bash
pytest -q
```

Sie prüfen Spaltenerkennung, Gruppierung, 16-Kanal-Aufteilung, Blockkopie samt Style/Merge/Maßen sowie den Schutz vor Überschreiben der Master- und Ausgabedatei.

## Linux-Binary und Windows-EXE erstellen

Ein eigenständiges Linux-Programm kann direkt unter Linux gebaut werden:

```bash
source .venv/bin/activate
pyinstaller --clean --noconfirm SmartStruxure_Beschriftungsgenerator.spec
```

Das Linux-Programm liegt dann unter `dist/SmartStruxure_Beschriftungsgenerator`. Tk/Tcl muss vor dem Build korrekt installiert sein, sonst kann PyInstaller die GUI nicht bündeln.

PyInstaller erzeugt Programme für das Betriebssystem, auf dem es läuft. Der Linux-Build ist daher **keine Windows-EXE**. Eine native Windows-EXE wird auf Windows gebaut. Das Repository dort auschecken und ausführen:

```powershell
py -3.12 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
pyinstaller --clean --noconfirm SmartStruxure_Beschriftungsgenerator.spec
```

Die EXE liegt danach unter:

```text
dist/SmartStruxure_Beschriftungsgenerator.exe
```

Die Spec-Datei bündelt `template_config.json` und – sofern vorhanden – die Mastervorlage. Für später leicht editierbare Konfigurationen kann `template_config.json` zusätzlich neben die EXE gelegt werden; diese externe Datei hat Vorrang. Dasselbe gilt für `templates/Beschriftung SmartStruxure ERR.xlsx` relativ zum EXE-Ordner.

Alternativ erzeugt der enthaltene GitHub-Actions-Workflow auf einem Windows-Runner automatisch die `.exe`. Wine-basierte Cross-Builds sind möglich, aber deutlich fehleranfälliger und für dieses Projekt nicht empfohlen.

## GitHub-Repository anlegen

Nach dem Prüfen lokaler Dateien:

```bash
git init
git add .
git commit -m "Initial SmartStruxure generator"
git branch -M main
git remote add origin https://github.com/DEIN-NAME/SmartStruxureGenerator.git
git push -u origin main
```

Die Referenzdatei kann interne Projektdaten enthalten. Vor einem öffentlichen Push sollte geprüft werden, ob sie veröffentlicht werden darf. Falls nicht, `templates/*.xlsx` zur `.gitignore` hinzufügen und die Vorlage getrennt verteilen.

## Projektstruktur

```text
main.py                         Programmeinstieg
cli.py                          Linux-/Kommandozeilentest ohne GUI
gui.py                          Tkinter-Oberfläche, Vorschau und Mappingdialog
excel_reader.py                 CSV-/Excel-Einlesen und Sheet-Auswahl
mapping.py                      Alias-Erkennung, Normalisierung, Gruppierung
template_analyzer.py            Visuelle Analyse und Template-Auswahl
generator.py                    Blockkopie, Befüllung und Ausgabe
models.py                       Datenmodelle
config.py                       Pfade, Defaults und Logging
template_config.json            Korrigierbare fachliche Zellzuordnung
inspect_template.py             Entwicklerwerkzeug
create_example_input.py         Erzeugt die Beispiel-Rohdatei
tests/                          Automatische Kernlogiktests
SmartStruxure_...spec           PyInstaller-Buildbeschreibung
```

## Bekannte Grenzen

- `openpyxl` bewahrt Zellformatierung, Merges, Maße und normale Excel-Inhalte. Excel-Objekte wie ActiveX-Steuerelemente, manche Zeichnungen oder externe Verbindungen sind nicht Teil der kopierten Kästchenlogik.
- `.xlsm` wird als Rohdatenquelle unterstützt. Die erzeugte Datei ist absichtlich `.xlsx` und enthält keine Makros.
