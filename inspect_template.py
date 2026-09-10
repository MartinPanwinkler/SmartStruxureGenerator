from __future__ import annotations

import argparse
import json
from pathlib import Path

from template_analyzer import analyze_template


def main() -> None:
    parser = argparse.ArgumentParser(description="Analysiert eine Excel-Mastervorlage ohne sie zu verändern.")
    parser.add_argument("template", type=Path)
    parser.add_argument("--json", type=Path, dest="json_output", help="Vollständigen Bericht als JSON speichern")
    args = parser.parse_args()
    report = analyze_template(args.template)
    for sheet_name, sheet in report["sheets"].items():
        print(f"\nSheet: {sheet_name}")
        print(f"Dimension: {sheet['dimensions']}")
        print("\nCandidate boxes:")
        for value in sheet["candidate_boxes"]:
            print(value)
        print("\nMerged Cells:")
        for value in sheet["merged_cells"]:
            print(value)
        print("\nRows:")
        for row, height in sheet["row_heights"].items():
            print(f"{row} height={height}")
        print("\nColumns:")
        for column, width in sheet["column_widths"].items():
            print(f"{column} width={width}")
        print("\nCells with values:")
        for cell in sheet["cells_with_values"]:
            print(f"{cell['cell']} = {cell['value']}")
    if args.json_output:
        args.json_output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"\nJSON-Bericht: {args.json_output}")


if __name__ == "__main__":
    main()
