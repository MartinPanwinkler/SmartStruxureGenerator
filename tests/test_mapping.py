from mapping import detect_columns, group_records, normalize_data, split_records


def test_detects_german_and_english_aliases() -> None:
    result = detect_columns(["ASP", "Modul", "Typ", "Channel", "DP-Name", "Beschreibung"])
    assert result["module_type"] == "Typ"
    assert result["channel"] == "Channel"
    assert result["datapoint"] == "DP-Name"


def test_grouping_uses_configured_fields() -> None:
    rows = [
        {"ASP": "A", "Modul": "1", "Typ": "DI", "Kanal": "1"},
        {"ASP": "A", "Modul": "1", "Typ": "DI", "Kanal": "2"},
        {"ASP": "A", "Modul": "2", "Typ": "DI", "Kanal": "1"},
    ]
    records = normalize_data(rows, {"asp": "ASP", "module": "Modul", "module_type": "Typ", "channel": "Kanal"})
    groups = group_records(records, ["asp", "module", "module_type"])
    assert [len(group) for _, group in groups] == [2, 1]


def test_splits_at_sixteen_channels() -> None:
    records = normalize_data([{"Kanal": str(i)} for i in range(24)], {"channel": "Kanal"})
    chunks = split_records(records, 16)
    assert [len(chunk) for chunk in chunks] == [16, 8]
