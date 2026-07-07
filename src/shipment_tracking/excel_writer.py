from __future__ import annotations

from datetime import datetime
import gc
import json
import os
from pathlib import Path
import subprocess
import tempfile
import time
from typing import Iterable
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from .models import TrackingRecord
from .time_utils import date_only


COL_BILL_NO = "\u63d0\u5355\u53f7"
COL_FORWARDER = "\u8d27\u4ee3"
COL_STATUS = "\u72b6\u6001\u663e\u793a"
COL_ARRIVAL = "\u5230\u6e2f\u65e5\u671f"
COL_EXCEPTION = "\u5f02\u5e38\u8bb0\u5f55"

COL_BILL_NO_EN = "BILL_NO"
COL_FORWARDER_EN = "GOODS_YARD"
COL_STATUS_EN = "DISPLAY_STATUS"
COL_ARRIVAL_EN = "ARRIVAL_DATE"

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
    records = list(records)
    remarks = remarks or {}
    if os.name == "nt" and os.getenv("SHIPMENT_TRACKING_USE_EXCEL_COM") == "1":
        try:
            return _update_tracking_workbook_with_excel_com(source_path, records, output_path, sheet_name, remarks)
        except RuntimeError:
            pass
    return _update_tracking_workbook_with_openpyxl(source_path, records, output_path, sheet_name, remarks)


def _update_tracking_workbook_with_openpyxl(
    source_path: str | Path,
    records: list[TrackingRecord],
    output_path: str | Path,
    sheet_name: str | None,
    remarks: dict[str, str],
) -> int:
    source = Path(source_path)
    output = Path(output_path)
    preserve_source = source
    temp_preserve: Path | None = None
    if source.resolve() == output.resolve():
        with tempfile.NamedTemporaryFile(delete=False, suffix=source.suffix) as temp_file:
            temp_file.write(source.read_bytes())
            temp_preserve = Path(temp_file.name)
            preserve_source = temp_preserve

    workbook = load_workbook(preserve_source)
    sheet = workbook[sheet_name] if sheet_name else workbook.active
    header_row, headers = _find_header_row(sheet)
    columns = _column_map(headers)

    key_col = _require_column(columns, [COL_BILL_NO, COL_BILL_NO_EN, "\u8fd0\u5355", "bl number", "b/l", "house bill", "tracking_number"])
    carrier_col = _find_column(columns, [COL_FORWARDER, COL_FORWARDER_EN, "forwarder", "carrier"])
    arrival_col = _require_column(columns, [COL_ARRIVAL, COL_ARRIVAL_EN])
    exception_col = _require_column(columns, [COL_EXCEPTION])

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
                row_changed |= _set_if_changed(
                    sheet.cell(row=row, column=arrival_col),
                    date_only(record.arrival_date),
                    date_format="m/d/yyyy",
                )

        if remark:
            cell = sheet.cell(row=row, column=exception_col)
            cell.value = _append_text(cell.value, remark)
            if remark.startswith("ERROR") or remark.startswith("NOT_FOUND"):
                cell.fill = FILL_ERROR
            elif "NO_ARRIVAL_DATE" in remark:
                cell.fill = FILL_WARNING

        if row_changed or remark:
            updated += 1

    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)
    workbook.close()
    _restore_unmanaged_package_parts(preserve_source, output)
    gc.collect()
    if temp_preserve and temp_preserve.exists():
        _unlink_with_retry(temp_preserve, raise_on_failure=False)
    return updated


def _update_tracking_workbook_with_excel_com(
    source_path: str | Path,
    records: list[TrackingRecord],
    output_path: str | Path,
    sheet_name: str | None,
    remarks: dict[str, str],
) -> int:
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    payload = _excel_com_payload(records, remarks)

    with tempfile.TemporaryDirectory(prefix="shipment_tracking_excel_") as temp_dir:
        temp_path = Path(temp_dir)
        payload_path = temp_path / "payload.json"
        script_path = temp_path / "update_excel.ps1"
        payload_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        script_path.write_text(_EXCEL_COM_SCRIPT, encoding="utf-8")

        completed = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
                str(source),
                str(output),
                sheet_name or "",
                str(payload_path),
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
        )

    if completed.returncode != 0:
        details = (completed.stderr or completed.stdout).strip()
        raise RuntimeError(f"Excel COM update failed: {details}")

    for line in completed.stdout.splitlines():
        if line.startswith("UPDATED_ROWS="):
            return int(line.split("=", 1)[1])
    raise RuntimeError(f"Excel COM update did not report updated rows: {completed.stdout.strip()}")


def _excel_com_payload(records: list[TrackingRecord], remarks: dict[str, str]) -> list[dict[str, str | bool | None]]:
    record_map = {record.tracking_number: record for record in records}
    tracking_numbers = sorted(set(record_map) | set(remarks))
    payload: list[dict[str, str | bool | None]] = []
    for tracking_number in tracking_numbers:
        record = record_map.get(tracking_number)
        payload.append(
            {
                "tracking_number": tracking_number,
                "carrier": record.carrier if record else None,
                "found": record.found if record else False,
                "arrival_date": _display_date(record.arrival_date) if record and record.arrival_date else None,
                "remark": remarks.get(tracking_number),
            }
        )
    return payload


def _display_date(value: datetime) -> str:
    return f"{value.month}/{value.day}/{value.year}"


def _restore_unmanaged_package_parts(source_path: Path, output_path: Path) -> None:
    if source_path.resolve() == output_path.resolve():
        original_bytes = source_path.read_bytes()
        with tempfile.NamedTemporaryFile(delete=False, suffix=source_path.suffix) as temp_original:
            temp_original.write(original_bytes)
            temp_original_path = Path(temp_original.name)
    else:
        temp_original_path = source_path

    temp_output = output_path.with_suffix(f"{output_path.suffix}.tmp")
    try:
        with ZipFile(temp_original_path) as source_zip, ZipFile(output_path) as output_zip:
            source_names = set(source_zip.namelist())
            output_names = set(output_zip.namelist())
            preserved_names = {
                name
                for name in source_names
                if name.startswith("customXml/") or name.startswith("[trash]/")
            }

            with ZipFile(temp_output, "w", ZIP_DEFLATED) as merged_zip:
                for item in output_zip.infolist():
                    data = output_zip.read(item.filename)
                    if item.filename == "[Content_Types].xml" and item.filename in source_names:
                        data = _merge_content_types(source_zip.read(item.filename), data)
                    elif item.filename == "xl/_rels/workbook.xml.rels" and item.filename in source_names:
                        data = _merge_workbook_relationships(source_zip.read(item.filename), data)
                    merged_zip.writestr(item, data)

                for name in sorted(preserved_names - output_names):
                    merged_zip.writestr(name, source_zip.read(name))

        _replace_with_retry(temp_output, output_path)
    finally:
        if source_path.resolve() == output_path.resolve() and temp_original_path.exists():
            temp_original_path.unlink()
        if temp_output.exists():
            temp_output.unlink()


def _merge_content_types(source_xml: bytes, output_xml: bytes) -> bytes:
    namespace = "http://schemas.openxmlformats.org/package/2006/content-types"
    ET.register_namespace("", namespace)
    source_root = ET.fromstring(source_xml)
    output_root = ET.fromstring(output_xml)
    existing_defaults = {element.attrib.get("Extension") for element in output_root.findall(f"{{{namespace}}}Default")}
    existing_overrides = {element.attrib.get("PartName") for element in output_root.findall(f"{{{namespace}}}Override")}

    for element in source_root:
        if element.tag == f"{{{namespace}}}Default":
            key = element.attrib.get("Extension")
            if key not in existing_defaults:
                output_root.append(element)
                existing_defaults.add(key)
        elif element.tag == f"{{{namespace}}}Override":
            key = element.attrib.get("PartName")
            if key and (key.startswith("/customXml/") or key.startswith("/[trash]/")) and key not in existing_overrides:
                output_root.append(element)
                existing_overrides.add(key)

    return ET.tostring(output_root, encoding="utf-8", xml_declaration=True)


def _replace_with_retry(source: Path, target: Path) -> None:
    last_error: PermissionError | None = None
    for _ in range(10):
        try:
            _unlink_with_retry(target)
            source.replace(target)
            return
        except PermissionError as exc:
            last_error = exc
            time.sleep(0.5)
    if last_error:
        raise last_error


def _unlink_with_retry(path: Path, *, raise_on_failure: bool = True) -> None:
    if not path.exists():
        return
    last_error: PermissionError | None = None
    for _ in range(10):
        try:
            path.unlink()
            return
        except PermissionError as exc:
            last_error = exc
            time.sleep(0.5)
    if last_error and raise_on_failure:
        raise last_error


def _merge_workbook_relationships(source_xml: bytes, output_xml: bytes) -> bytes:
    namespace = "http://schemas.openxmlformats.org/package/2006/relationships"
    ET.register_namespace("", namespace)
    source_root = ET.fromstring(source_xml)
    output_root = ET.fromstring(output_xml)
    existing_targets = {element.attrib.get("Target") for element in output_root.findall(f"{{{namespace}}}Relationship")}

    for element in source_root.findall(f"{{{namespace}}}Relationship"):
        rel_type = element.attrib.get("Type", "")
        target = element.attrib.get("Target")
        if rel_type.endswith("/customXml") and target not in existing_targets:
            output_root.append(element)
            existing_targets.add(target)

    return ET.tostring(output_root, encoding="utf-8", xml_declaration=True)


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


def _append_text(existing, addition: str) -> str:
    current = str(existing).strip() if existing is not None else ""
    if not current:
        return addition
    return f"{current}\n{addition}"


def _find_header_row(sheet) -> tuple[int, list[str]]:
    for required_set in ({COL_BILL_NO, COL_FORWARDER, COL_STATUS}, {COL_BILL_NO_EN, COL_FORWARDER_EN, COL_STATUS_EN}):
        for row_number in range(1, min(sheet.max_row, 20) + 1):
            headers = [str(cell.value).strip() if cell.value is not None else "" for cell in sheet[row_number]]
            lowered = [header.lower() for header in headers]
            if all(required.lower() in lowered for required in required_set):
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


_EXCEL_COM_SCRIPT = r'''
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$OutputPath,
    [Parameter(Mandatory=$false)][string]$SheetName,
    [Parameter(Mandatory=$true)][string]$PayloadPath
)

$ErrorActionPreference = 'Stop'
$payload = Get-Content -Path $PayloadPath -Raw -Encoding UTF8 | ConvertFrom-Json
$items = @{}
foreach ($item in $payload) {
    $items[[string]$item.tracking_number] = $item
}

function Normalize([object]$value) {
    if ($null -eq $value) { return '' }
    return ([string]$value).Trim().ToLowerInvariant()
}

function Find-Column($headers, [string[]]$names, [bool]$allowContains) {
    foreach ($name in $names) {
        $needle = Normalize $name
        foreach ($key in $headers.Keys) {
            if ($key -eq $needle) { return $headers[$key] }
        }
    }
    if ($allowContains) {
        foreach ($name in $names) {
            $needle = Normalize $name
            foreach ($key in $headers.Keys) {
                if ($key.Contains($needle)) { return $headers[$key] }
            }
        }
    }
    return $null
}

function Find-Header($sheet, [int]$maxCol) {
    $requiredSets = @(
        @('提单号', '货代', '状态显示'),
        @('BILL_NO', 'GOODS_YARD', 'DISPLAY_STATUS')
    )
    foreach ($requiredSet in $requiredSets) {
        for ($row = 1; $row -le 20; $row++) {
            $headers = @{}
            for ($col = 1; $col -le $maxCol; $col++) {
                $value = $sheet.Cells.Item($row, $col).Text
                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    $headers[(Normalize $value)] = $col
                }
            }
            $ok = $true
            foreach ($required in $requiredSet) {
                if (-not $headers.ContainsKey((Normalize $required))) {
                    $ok = $false
                    break
                }
            }
            if ($ok) {
                return @{ Row = $row; Headers = $headers }
            }
        }
    }
    throw 'Could not find expected header row.'
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null
$updated = 0

try {
    $workbook = $excel.Workbooks.Open($WorkbookPath)
    if ([string]::IsNullOrWhiteSpace($SheetName)) {
        $sheet = $workbook.Worksheets.Item(1)
    } else {
        $sheet = $workbook.Worksheets.Item($SheetName)
    }

    $used = $sheet.UsedRange
    $maxRow = $used.Row + $used.Rows.Count - 1
    $maxCol = $used.Column + $used.Columns.Count - 1
    $headerInfo = Find-Header $sheet $maxCol
    $headerRow = [int]$headerInfo.Row
    $headers = $headerInfo.Headers

    $keyCol = Find-Column $headers @('提单号', 'BILL_NO', '运单', 'bl number', 'b/l', 'house bill', 'tracking_number') $true
    $carrierCol = Find-Column $headers @('货代', 'GOODS_YARD', 'forwarder', 'carrier') $true
    $arrivalCol = Find-Column $headers @('到港日期', 'ARRIVAL_DATE') $false
    $exceptionCol = Find-Column $headers @('异常记录') $false

    if ($null -eq $keyCol -or $null -eq $arrivalCol) {
        throw 'Missing required 提单号/BILL_NO or 到港日期/ARRIVAL_DATE column.'
    }
    if ($null -eq $exceptionCol) {
        throw 'Missing required 异常记录 column.'
    }

    for ($row = $headerRow + 1; $row -le $maxRow; $row++) {
        $trackingNumber = ([string]$sheet.Cells.Item($row, $keyCol).Text).Trim()
        if ([string]::IsNullOrWhiteSpace($trackingNumber) -or -not $items.ContainsKey($trackingNumber)) {
            continue
        }
        $item = $items[$trackingNumber]
        if ($null -ne $item.carrier -and $null -ne $carrierCol) {
            $carrier = ([string]$sheet.Cells.Item($row, $carrierCol).Text).Trim().ToUpperInvariant()
            if (-not $carrier.Contains(([string]$item.carrier).ToUpperInvariant())) {
                continue
            }
        }

        $rowChanged = $false
        if ($item.found -and -not [string]::IsNullOrWhiteSpace([string]$item.arrival_date)) {
            $arrivalCell = $sheet.Cells.Item($row, $arrivalCol)
            $arrivalDate = [datetime]::ParseExact([string]$item.arrival_date, 'M/d/yyyy', [Globalization.CultureInfo]::InvariantCulture)
            $arrivalCell.Value2 = $arrivalDate.ToOADate()
            $arrivalCell.NumberFormat = 'm/d/yyyy'
            $arrivalCell.Interior.Color = 13431551
            $rowChanged = $true
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$item.remark)) {
            $remarkCell = $sheet.Cells.Item($row, $exceptionCol)
            $remark = [string]$item.remark
            $existing = ([string]$remarkCell.Text).Trim()
            if ([string]::IsNullOrWhiteSpace($existing)) {
                $remarkCell.Value2 = $remark
            } else {
                $remarkCell.Value2 = $existing + [Environment]::NewLine + $remark
            }
            if ($remark.StartsWith('ERROR') -or $remark.StartsWith('NOT_FOUND')) {
                $remarkCell.Interior.Color = 13421823
            } elseif ($remark.Contains('NO_ARRIVAL_DATE')) {
                $remarkCell.Interior.Color = 13495295
            }
            $rowChanged = $true
        }

        if ($rowChanged) {
            $updated++
        }
    }

    if ([IO.Path]::GetFullPath($WorkbookPath) -eq [IO.Path]::GetFullPath($OutputPath)) {
        $workbook.Save()
    } else {
        $workbook.SaveAs($OutputPath)
    }
    Write-Output "UPDATED_ROWS=$updated"
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        [Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    $excel.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
'''
