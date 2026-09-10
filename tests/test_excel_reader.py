from pathlib import Path

from openpyxl import Workbook

from excel_reader import load_source_file
from mapping import detect_columns, group_records, normalize_data


def test_smartstruxure_export_is_parsed_without_column_mapping(tmp_path: Path) -> None:
    path = tmp_path / "export.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "ASP-999_DP"
    sheet.append([None, "RTD-DI-16", "M3 RTD-DI-16", "RTD-DI-16", 3, None, None])
    sheet.append([None, "2-Kabel RTD", "DP.ONE", "Temperatur eins", None, "In1", "Quelle 1"])
    sheet.append([None, "2-Kabel RTD", "DP.TWO", "Temperatur zwei", None, "In2", "Quelle 2"])
    workbook.save(path)

    table = load_source_file(path)
    mapping = detect_columns(table.headers)
    records = normalize_data(table.rows, mapping)

    assert len(records) == 2
    assert records[0].asp == "ASP-999"
    assert records[0].module == "M3 RTD-DI-16"
    assert records[0].module_type == "RTD-DI-16"
    assert records[0].channel == "In1"
    assert records[1].datapoint == "DP.TWO"
    assert len(group_records(records, ["asp", "module", "module_type"])) == 1
