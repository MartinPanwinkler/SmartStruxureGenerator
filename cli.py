from __future__ import annotations

import argparse
from pathlib import Path

from config import DEFAULT_CONFIG, DEFAULT_TEMPLATE, load_json_config, setup_logging
from excel_reader import load_source_file
from generator import generate_workbook
from mapping import detect_columns, normalize_data


def main() -> None:
    parser = argparse.ArgumentParser(description="SmartStruxure-Generator ohne grafische Oberfläche")
    parser.add_argument("source", type=Path, help="Rohdatei (.xlsx, .xlsm oder .csv)")
    parser.add_argument("output", type=Path, help="Neue .xlsx-Ausgabedatei")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE, help="Master-Vorlage")
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG, help="template_config.json")
    parser.add_argument("--sheet", help="Tabellenblatt der Rohdatei")
    args = parser.parse_args()
    setup_logging()
    source = load_source_file(args.source, args.sheet)
    mapping = detect_columns(source.headers)
    required = ("module", "module_type", "channel", "description")
    missing = [field for field in required if field not in mapping]
    if missing:
        parser.error("Nicht automatisch erkannte Pflichtfelder: " + ", ".join(missing))
    records = normalize_data(source.rows, mapping)
    result = generate_workbook(args.template, args.output, records, load_json_config(args.config))
    print(f"Fertig: {result.boxes_created} Kästchen, {result.records_processed} Datenpunkte")
    print(result.output_path)
    for warning in result.warnings:
        print(f"Warnung: {warning}")


if __name__ == "__main__":
    main()
