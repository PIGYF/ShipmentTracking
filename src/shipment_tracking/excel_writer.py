from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from .models import TrackingRecord
from .time_utils import date_only


COL_BILL_NO = "\u63d0\u5355\u53f7"
COL_FORWARDER = "\u8d27\u4ee3"
COL_STATUS = "\u72b6\u6001\u663e\u793a"
COL_ARRIVAL = "\u5230\u6e2f\u65e5\u671f"
COL_ARRIVAL_TYPE = "\u5230\u6e2f\u65e5\u671f\u7c7b\u578b"
COL_REMARK = "\u67e5\u8be2\u5907\u6ce8"

FILL_CHANGED = PatternFill("solid", fgColor="FFF2CC")
FILL_ERROR = PatternFill("solid", fgColor="F4CCCC")
FILL_WARNING = PatternFill("solid", fgColor="FCE5CD")


def update_tracking_workbook(
    source_path: str | Path,
    records: Iterable[TrackingRecord],
    output_path: str | Path,
    sheet_name: str | None = None,
    remarks: dict[str, str] | None = None,
) -> int:
    workbook = load_workbook(Path(source_path))
    sheet = workbook[sheet_name] if sheet_name else workbook.active
    header_row, headers = _find_header_row(sheet)
    columns = _column_map(headers)

    key_col = _require_column(columns, [COL_BILL_NO, "\u8fd0\u5355", "bl number", "b/l", "house bill", "tracking_number"])
    carrier_col = _find_column(columns, [COL_FORWARDER, "forwarder", "carrier"])
    arrival_col = _ensure_column(sheet, header_row, columns, COL_ARRIVAL)
    status_col = _ensure_column(sheet, header_row, columns, COL_STATUS)
    arrival_type_col = _ensure_column(sheet, header_row, columns, COL_ARRIVAL_TYPE)
    remark_col = _ensure_column(sheet, header_row, columns, COL_REMARK)

    record_map = {record.tracking_number: record for record in records}
    remarks = remarks or {}
    updated = 0

    for row in range(header_row + 1, sheet.max_row + 1):
        tracking_number = str(sheet.cell(row=row, column=key_col).value or "").strip()
        record = record_map.get(tracking_number)
        remark = remarks.get(tracking_number)
        if not record and not remark:
            continue

        if record and carrier_col:
            carrier = str(sheet.cell(row=row, column=carrier_col).value or "").strip().upper()
            if record.carrier.upper() not in carrier:
                continue

        row_changed = False
        if record and record.found:
            if record.arrival_date:
                row_changed |= _set_if_changed(sheet.cell(row=row, column=arrival_col), date_only(record.arrival_date), date_format="yyyy-mm-dd")
            if record.status_description or record.status:
                row_changed |= _set_if_changed(sheet.cell(row=row, column=status_col), record.status_description or record.status)
            if record.arrival_date_type:
                row_changed |= _set_if_changed(sheet.cell(row=row, column=arrival_type_col), record.arrival_date_type)

        if remark:
            cell = sheet.cell(row=row, column=remark_col)
            cell.value = remark
            if remark.startswith("ERROR") or remark.startswith("NOT_FOUND"):
                cell.fill = FILL_ERROR
            elif "NO_ARRIVAL_DATE" in remark:
                cell.fill = FILL_WARNING

        if row_changed or remark:
            updated += 1

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return updated


def _set_if_changed(cell, new_value, date_format: str | None = None) -> bool:
    if new_value is None:
        return False
    old_value = _normalize_cell_value(cell.value)
    comparable = _normalize_cell_value(new_value)
    if old_value == comparable:
        return False
    cell.value = new_value
    if date_format:
        cell.number_format = date_format
    cell.fill = FILL_CHANGED
    return True


def _normalize_cell_value(value):
    if isinstance(value, datetime):
        return value.replace(hour=0, minute=0, second=0, microsecond=0)
    return value


def _find_header_row(sheet) -> tuple[int, list[str]]:
    aliases = {COL_BILL_NO, "\u8fd0\u5355", COL_FORWARDER, "bl number", "b/l", "house bill", "forwarder", "carrier"}
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
