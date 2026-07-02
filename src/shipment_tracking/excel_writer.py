from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook

from .models import TrackingRecord


def update_tracking_workbook(
    source_path: str | Path,
    records: Iterable[TrackingRecord],
    output_path: str | Path,
    sheet_name: str | None = None,
) -> int:
    workbook = load_workbook(Path(source_path))
    sheet = workbook[sheet_name] if sheet_name else workbook.active
    header_row, headers = _find_header_row(sheet)
    columns = _column_map(headers)

    key_col = _require_column(columns, ["提单号", "运单", "bl number", "b/l", "house bill", "tracking_number"])
    carrier_col = _find_column(columns, ["货代", "forwarder", "carrier"])
    arrival_col = _ensure_column(sheet, header_row, columns, "到港日期")
    status_col = _ensure_column(sheet, header_row, columns, "状态显示")
    arrival_type_col = _ensure_column(sheet, header_row, columns, "到港日期类型")

    record_map = {record.tracking_number: record for record in records if record.found}
    updated = 0
    for row in range(header_row + 1, sheet.max_row + 1):
        tracking_number = str(sheet.cell(row=row, column=key_col).value or "").strip()
        record = record_map.get(tracking_number)
        if not record:
            continue
        if carrier_col:
            carrier = str(sheet.cell(row=row, column=carrier_col).value or "").strip().upper()
            if record.carrier.upper() not in carrier:
                continue

        if record.arrival_date:
            sheet.cell(row=row, column=arrival_col).value = _excel_datetime(record.arrival_date)
        if record.status_description or record.status:
            sheet.cell(row=row, column=status_col).value = record.status_description or record.status
        if record.arrival_date_type:
            sheet.cell(row=row, column=arrival_type_col).value = record.arrival_date_type
        updated += 1

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return updated


def _find_header_row(sheet) -> tuple[int, list[str]]:
    aliases = {"提单号", "运单", "货代", "bl number", "b/l", "house bill", "forwarder", "carrier"}
    for row_number in range(1, min(sheet.max_row, 20) + 1):
        headers = [str(cell.value).strip() if cell.value is not None else "" for cell in sheet[row_number]]
        lowered = [header.lower() for header in headers]
        if any(alias.lower() in lowered for alias in aliases):
            return row_number, headers
    raise ValueError("Could not find a shipment header row in the first 20 rows.")


def _column_map(headers: list[str]) -> dict[str, int]:
    return {header.lower(): index + 1 for index, header in enumerate(headers) if header}


def _find_column(columns: dict[str, int], names: list[str]) -> int | None:
    for name in names:
        needle = name.lower()
        for header, column in columns.items():
            if header == needle or needle in header:
                return column
    return None


def _require_column(columns: dict[str, int], names: list[str]) -> int:
    column = _find_column(columns, names)
    if not column:
        raise ValueError(f"Missing required column. Expected one of: {', '.join(names)}")
    return column


def _ensure_column(sheet, header_row: int, columns: dict[str, int], header: str) -> int:
    existing = _find_column(columns, [header])
    if existing:
        return existing
    column = sheet.max_column + 1
    sheet.cell(row=header_row, column=column).value = header
    columns[header.lower()] = column
    return column


def _excel_datetime(value: datetime) -> datetime:
    if value.tzinfo:
        value = value.replace(tzinfo=None)
    return value.replace(hour=0, minute=0, second=0, microsecond=0)
