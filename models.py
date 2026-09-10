from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable


@dataclass(slots=True)
class SourceTable:
    headers: list[str]
    rows: list[dict[str, Any]]
    sheets: list[str] = field(default_factory=list)
    selected_sheet: str | None = None


@dataclass(slots=True)
class Record:
    asp: str = ""
    automation_server: str = ""
    module: str = ""
    module_type: str = ""
    module_number: str = ""
    channel: str = ""
    datapoint: str = ""
    description: str = ""
    source: str = ""
    target: str = ""
    address: str = ""
    extra: dict[str, Any] = field(default_factory=dict)

    def value(self, key: str) -> Any:
        return getattr(self, key, self.extra.get(key, ""))


@dataclass(slots=True)
class BoxTemplate:
    name: str
    sheet: str
    min_row: int
    max_row: int
    min_col: int
    max_col: int
    header_cells: dict[str, str | list[str]] = field(default_factory=dict)
    data_start_row: int = 1
    max_channels: int = 1
    columns: dict[str, str] = field(default_factory=dict)
    data_slots: list[dict[str, str]] = field(default_factory=list)
    header_formats: dict[str, str] = field(default_factory=dict)
    preserve_template_fields: set[str] = field(default_factory=set)
    merged_cells: list[str] = field(default_factory=list)
    row_heights: dict[int, float | None] = field(default_factory=dict)
    column_widths: dict[int, float | None] = field(default_factory=dict)

    @property
    def height(self) -> int:
        return self.max_row - self.min_row + 1

    @property
    def width(self) -> int:
        return self.max_col - self.min_col + 1


@dataclass(slots=True)
class GenerationResult:
    output_path: Path
    boxes_created: int
    records_processed: int
    warnings: list[str] = field(default_factory=list)


ProgressCallback = Callable[[int, str], None]
