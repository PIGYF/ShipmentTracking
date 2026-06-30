from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook


@dataclass(frozen=True)
class ShipmentInput:
    row_number: int
    tracking_number: str
    forwarder: str


def read_shipments_from_excel(path: str | Path) -> list[ShipmentInput]:
    workbook = load_workbook(Path(path), read_only=True, data_only=True)
    sheet = workbook.active
    headers = [str(cell.value).strip() if cell.value is not None else "" for cell in sheet[1]]

    tracking_idx = _find_column(headers, ["运单", "bl number", "house bill", "tracking_number"])
    forwarder_idx = _find_column(headers, ["货代", "forwarder", "carrier"])

    rows: list[ShipmentInput] = []
    for row_number, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        tracking_value = row[tracking_idx] if tracking_idx < len(row) else None
        forwarder_value = row[forwarder_idx] if forwarder_idx < len(row) else None
        if tracking_value is None:
            continue
        rows.append(
            ShipmentInput(
                row_number=row_number,
                tracking_number=str(tracking_value).strip(),
                forwarder=str(forwarder_value or "").strip(),
            )
        )
    return rows


def _find_column(headers: list[str], names: list[str]) -> int:
    lowered = [header.lower() for header in headers]
    for name in names:
        if name.lower() in lowered:
            return lowered.index(name.lower())
    raise ValueError(f"Missing required column. Expected one of: {', '.join(names)}")
