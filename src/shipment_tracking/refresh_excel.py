from __future__ import annotations

import argparse
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
import socket
import sys
import time
from urllib.error import URLError

from openpyxl import load_workbook

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from shipment_tracking.dgf import DgfClient
from shipment_tracking.env import load_dotenv
from shipment_tracking.excel_writer import COL_BILL_NO, COL_FORWARDER, COL_STATUS, update_tracking_workbook
from shipment_tracking.maersk import MaerskClient
from shipment_tracking.models import TrackingRecord


DEFAULT_SHEET = "2026"
PENDING_STATUS = "\u672a\u9001\u8d27"
LOG_FILENAME = "refresh.log"
ENABLED_CARRIERS = {"DGF", "MAERSK"}
_log_path: Path | None = None


def main() -> None:
    load_dotenv()
    parser = argparse.ArgumentParser(description="Refresh supported shipment tracking data in an Excel workbook.")
    parser.add_argument("excel_path", help="Input workbook path.")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help="Worksheet name. Default: 2026.")
    parser.add_argument("--output", help="Output workbook path. Default: overwrite input workbook.")
    parser.add_argument("--limit", type=int, help="Limit unique API queries for testing.")
    parser.add_argument("--dry-run", action="store_true", help="Show what would be queried without calling APIs or writing Excel.")
    args = parser.parse_args()

    source = Path(args.excel_path)
    output = Path(args.output) if args.output else source
    _init_log(output)
    _emit(f"Workbook: {source}")
    _emit(f"Sheet: {args.sheet}")
    _emit(f"Output: {output}")
    _emit(f"Enabled carriers: {', '.join(sorted(ENABLED_CARRIERS))}")
    if args.dry_run:
        _emit("Mode: DRY-RUN")
    rows = _load_pending_supported_rows(source, args.sheet)
    unique_jobs = _unique_jobs(rows)
    if args.limit:
        unique_jobs = unique_jobs[: args.limit]

    counts = Counter(carrier for carrier, _ in unique_jobs)
    _emit(f"Loaded pending supported rows: {len(rows)}")
    _emit("Unique API queries: " + ", ".join(f"{carrier}={count}" for carrier, count in sorted(counts.items())) if counts else "Unique API queries: 0")

    if args.dry_run:
        for index, (carrier, tracking_number) in enumerate(unique_jobs, start=1):
            _emit(f"[DRY-RUN] {index}/{len(unique_jobs)} {carrier} {tracking_number}")
        return

    records, remarks = _run_jobs_by_carrier(unique_jobs)

    updated = update_tracking_workbook(source, records, output, args.sheet, remarks)
    _emit(f"Updated rows: {updated}")
    _emit(f"Saved: {output}")


def _run_jobs_by_carrier(unique_jobs: list[tuple[str, str]]) -> tuple[list[TrackingRecord], dict[str, str]]:
    grouped: dict[str, list[str]] = {carrier: [] for carrier in sorted(ENABLED_CARRIERS)}
    for carrier, tracking_number in unique_jobs:
        grouped.setdefault(carrier, []).append(tracking_number)

    active_groups = {carrier: jobs for carrier, jobs in grouped.items() if jobs}
    records: list[TrackingRecord] = []
    remarks: dict[str, str] = {}

    with ThreadPoolExecutor(max_workers=len(active_groups) or 1) as executor:
        futures = {
            executor.submit(_run_carrier_jobs, carrier, jobs): carrier
            for carrier, jobs in active_groups.items()
        }
        for future in as_completed(futures):
            carrier = futures[future]
            carrier_records, carrier_remarks = future.result()
            records.extend(carrier_records)
            remarks.update(carrier_remarks)
            total = len(active_groups[carrier])
            _emit(f"[{carrier}] completed {total}/{total}")

    return records, remarks


def _run_carrier_jobs(carrier: str, tracking_numbers: list[str]) -> tuple[list[TrackingRecord], dict[str, str]]:
    records: list[TrackingRecord] = []
    remarks: dict[str, str] = {}
    total = len(tracking_numbers)
    try:
        client = _client_for(carrier)
    except RuntimeError as exc:
        skipped = f"SKIPPED: {exc}"
        for tracking_number in tracking_numbers:
            remarks[tracking_number] = skipped
        _emit(f"[{carrier}] skipped {total}/{total}: {exc}")
        return records, remarks

    for index, tracking_number in enumerate(tracking_numbers, start=1):
        try:
            record = _track(client, carrier, tracking_number)
            records.append(record)
            remark = _remark(record) or _success_remark(record)
            if remark:
                remarks[tracking_number] = remark
            _emit(_progress_line(index, total, record, remark))
        except Exception as exc:
            error = _format_error(exc)
            remarks[tracking_number] = error
            _emit(f"[{carrier}] {index}/{total} {tracking_number} -> {error}")
            if _should_stop_carrier(exc):
                skipped = tracking_numbers[index:]
                for skipped_tracking_number in skipped:
                    remarks[skipped_tracking_number] = f"SKIPPED: {carrier} query stopped after connection/configuration error"
                if skipped:
                    _emit(
                        f"[{carrier}] stopped after API connection/configuration error; "
                        f"skipped remaining {len(skipped)}/{total}"
                    )
                break

        if carrier == "DGF" and index < total:
            time.sleep(6)

    return records, remarks


def _client_for(carrier: str):
    if carrier == "DGF":
        return DgfClient()
    if carrier == "MAERSK":
        return MaerskClient()
    raise ValueError(f"Unsupported carrier: {carrier}")


def _load_pending_supported_rows(path: Path, sheet_name: str) -> list[tuple[str, str]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook[sheet_name]
        header_row, headers = _find_header_row(sheet)
        columns = {header.lower(): index for index, header in enumerate(headers) if header}
        bill_idx = _find_column(columns, [COL_BILL_NO])
        carrier_idx = _find_column(columns, [COL_FORWARDER])
        status_idx = _find_column(columns, [COL_STATUS])

        rows: list[tuple[str, str]] = []
        for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
            status = str(row[status_idx] or "").strip()
            if status != PENDING_STATUS:
                continue
            carrier = str(row[carrier_idx] or "").strip().upper()
            if carrier not in ENABLED_CARRIERS:
                continue
            tracking_number = str(row[bill_idx] or "").strip()
            if tracking_number:
                rows.append((carrier, tracking_number))
        return rows
    finally:
        workbook.close()


def _find_header_row(sheet) -> tuple[int, list[str]]:
    aliases = {COL_BILL_NO, COL_FORWARDER, COL_STATUS}
    for row_number, row in enumerate(sheet.iter_rows(min_row=1, max_row=20, values_only=True), start=1):
        headers = [str(value).strip() if value is not None else "" for value in row]
        lowered = [header.lower() for header in headers]
        if all(alias.lower() in lowered for alias in aliases):
            return row_number, headers
    raise ValueError("Could not find expected header row with 提单号, 货代, 状态显示.")


def _unique_jobs(rows: list[tuple[str, str]]) -> list[tuple[str, str]]:
    seen: set[tuple[str, str]] = set()
    jobs: list[tuple[str, str]] = []
    for item in rows:
        if item in seen:
            continue
        seen.add(item)
        jobs.append(item)
    return jobs


def _find_column(columns: dict[str, int], names: list[str]) -> int:
    for name in names:
        needle = name.lower()
        for header, index in columns.items():
            if header == needle or needle in header:
                return index
    raise ValueError(f"Missing required column: {', '.join(names)}")


def _track(client, carrier: str, tracking_number: str) -> TrackingRecord:
    if carrier == "DGF":
        return client.track(tracking_number).to_record()
    if carrier == "MAERSK":
        return client.track(tracking_number).to_record()
    raise ValueError(f"Unsupported carrier: {carrier}")


def _remark(record: TrackingRecord) -> str | None:
    if not record.found:
        return "NOT_FOUND: API returned no shipment"
    if not record.arrival_date:
        return "NO_ARRIVAL_DATE: shipment found but no ETA/ATA"
    return None


def _success_remark(record: TrackingRecord) -> str:
    parts = [record.carrier, "FOUND"]
    if record.arrival_date:
        parts.append(f"{record.arrival_date_type or 'ARRIVAL'} {_format_date(record.arrival_date)}")
    if record.status_description:
        parts.append(record.status_description)
    elif record.status:
        parts.append(record.status)
    return ": ".join((parts[0], " | ".join(parts[1:])))


def _format_date(value) -> str:
    return f"{value.month}/{value.day}/{value.year}"


def _progress_line(index: int, total: int, record: TrackingRecord, remark: str | None) -> str:
    if remark:
        result = remark
    elif record.arrival_date:
        result = f"{record.arrival_date_type} {record.arrival_date:%Y-%m-%d %H:%M}"
    else:
        result = "OK"
    return f"[{record.carrier}] {index}/{total} {record.tracking_number} -> {result}"


def _format_error(exc: Exception) -> str:
    return f"ERROR: {type(exc).__name__}: {exc}"


def _should_stop_carrier(exc: Exception) -> bool:
    if isinstance(exc, RuntimeError):
        return True
    if isinstance(exc, (TimeoutError, ConnectionError, socket.timeout, URLError)):
        return True
    return False


def _init_log(excel_path: Path) -> None:
    global _log_path
    _log_path = excel_path.parent / LOG_FILENAME
    _log_path.write_text(f"Run started: {datetime.now():%Y-%m-%d %H:%M:%S}\n", encoding="utf-8")


def _emit(message: str) -> None:
    print(message, flush=True)
    if _log_path:
        with _log_path.open("a", encoding="utf-8") as log_file:
            log_file.write(f"{message}\n")


if __name__ == "__main__":
    main()
