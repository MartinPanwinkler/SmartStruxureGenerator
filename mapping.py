from __future__ import annotations

import re
import unicodedata
from collections.abc import Iterable
from typing import Any

from models import Record


COLUMN_ALIASES: dict[str, list[str]] = {
    "asp": ["ASP", "Automation Server", "AutomationServer", "AS"],
    "automation_server": ["Automation Server", "Server"],
    "module": ["Modul", "Module", "Baugruppe"],
    "module_type": ["Modultyp", "Modul Typ", "Module Type", "Typ", "IO-Typ"],
    "module_number": ["Modulnummer", "Module Number", "Modul Nr", "Steckplatz"],
    "channel": ["Kanal", "Channel", "Ein-/Ausgang", "IO", "I/O", "Eingang", "Ausgang"],
    "datapoint": ["Datenpunkt", "Datenpunktname", "DP-Name", "SmartStruxure", "SmartStruxure-Name"],
    "description": ["Beschreibung", "Bezeichnung", "Description", "Text"],
    "source": ["Quelle", "Source", "Von"],
    "target": ["Ziel", "Target", "Nach"],
    "address": ["Modicon-Adresse", "SPS-Adresse", "Adresse", "Address"],
}


def normalize_header(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]+", "", text.casefold())


def detect_columns(headers: Iterable[str], aliases: dict[str, list[str]] = COLUMN_ALIASES) -> dict[str, str]:
    actual = [str(header) for header in headers]
    normalized = {normalize_header(header): header for header in actual if normalize_header(header)}
    detected: dict[str, str] = {}
    used: set[str] = set()
    for internal, names in aliases.items():
        for candidate in [internal, *names]:
            match = normalized.get(normalize_header(candidate))
            if match and match not in used:
                detected[internal] = match
                used.add(match)
                break
    return detected


def normalize_data(rows: Iterable[dict[str, Any]], column_mapping: dict[str, str]) -> list[Record]:
    fields = set(Record.__dataclass_fields__) - {"extra"}
    records: list[Record] = []
    for row in rows:
        values = {
            internal: "" if row.get(source) is None else str(row.get(source)).strip()
            for internal, source in column_mapping.items()
            if internal in fields
        }
        extras = {key: value for key, value in row.items() if key not in column_mapping.values()}
        records.append(Record(**values, extra=extras))
    return records


def group_records(records: Iterable[Record], group_by: list[str]) -> list[tuple[tuple[str, ...], list[Record]]]:
    groups: dict[tuple[str, ...], list[Record]] = {}
    for record in records:
        key = tuple(str(record.value(field) or "") for field in group_by)
        groups.setdefault(key, []).append(record)
    return list(groups.items())


def split_records(records: list[Record], limit: int) -> list[list[Record]]:
    if limit <= 0:
        raise ValueError("max_channels muss größer als 0 sein.")
    return [records[index:index + limit] for index in range(0, len(records), limit)] or [[]]
