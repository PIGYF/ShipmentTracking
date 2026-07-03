from __future__ import annotations

import argparse
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
import sys
import time

from openpyxl import load_workbook

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from shipment_tracking.dgf import DgfClient
from shipment_tracking.dsv_public import DsvPublicClient
from shipment_tracking.env import load_dotenv
from shipment_tracking.excel_writer import COL_BILL_NO, COL_FORWARDER, COL_STATUS, update_tracking_workbook
from shipment_tracking.maersk import MaerskClient
from shipment_tracking.models import TrackingRecord


DEFAULT_SHEET = "2026"
PENDING_STATUS = "\u672a\u9001\u8d27"


def main() -> None:
    load_dotenv()
    parser = argparse.ArgumentParser(description="Refresh supported shipment tracking data in an Excel copy.")
    parser.add_argument("excel_path", help="Input workbook path.")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help="Worksheet name. Default: 2026.")
    parser.add_argument("--output", help="Output workbook path. Default: exports/<input>.updated.xlsx")
    parser.add_argument("--limit", type=int, help="Limit unique API queries for testing.")
    parser.add_argument("--dry-run", action="store_true", help="Show what would be queried without calling APIs or writing Excel.")
    args = parser.parse_args()

    source = Path(args.excel_path)
    output = Path(args.output) if args.output else Path("exports") / f"{source.stem}.updated{source.suffix}"
    rows = _load_pending_supported_rows(source, args.sheet)
    unique_jobs = _unique_jobs(rows)
    if args.limit:
        unique_jobs = unique_jobs[: args.limit]

    counts = Counter(carrier for carrier, _ in unique_jobs)
    print(f"Loaded pending supported rows: {len(rows)}")
    print("Unique API queries: " + ", ".join(f"{carrier}={count}" for carrier, count in sorted(counts.items())) if counts else "Unique API queries: 0")

    if args.dry_run:
        for index, (carrier, tracking_number) in enumerate(unique_jobs, start=1):
            print(f"[DRY-RUN] {index}/{len(unique_jobs)} {carrier} {tracking_number}")
        return

    records, remarks = _run_jobs_by_carrier(unique_jobs)

    updated = update_tracking_workbook(source, records, output, args.sheet, remarks)
    print(f"Updated rows: {updated}")
    print(f"Saved: {output}")


def _run_jobs_by_carrier(unique_jobs: list[tuple[str, str]]) -> tuple[list[TrackingRecord], dict[str, str]]:
    grouped: dict[str, list[str]] = {"DGF": [], "DSV": [], "MAERSK": []}
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
            print(f"[{carrier}] completed {total}/{total}")

    return records, remarks


def _run_carrier_jobs(carrier: str, tracking_numbers: list[str]) -> tuple[list[TrackingRecord], dict[str, str]]:
    client = _client_for(carrier)
    records: list[TrackingRecord] = []
    remarks: dict[str, str] = {}
    total = len(tracking_numbers)

    for index, tracking_number in enumerate(tracking_numbers, start=1):
        try:
            record = _track(client, carrier, tracking_number)
            records.append(record)
            remark = _remark(record)
            if remark:
                remarks[tracking_number] = remark
            print(_progress_line(index, total, record, remark), flush=True)
        except Exception as exc:
            remarks[tracking_number] = f"ERROR: {type(exc).__name__}: {exc}"
            print(f"[{carrier}] {index}/{total} {tracking_number} -> ERROR {type(exc).__name__}: {exc}", flush=True)

        if carrier == "DGF" and index < total:
            time.sleep(6)

    return records, remarks


def _client_for(carrier: str):
    if carrier == "DGF":
        return DgfClient()
    if carrier == "DSV":
        return DsvPublicClient()
    if carrier == "MAERSK":
        return MaerskClient()
    raise ValueError(f"Unsupported carrier: {carrier}")


def _load_pending_supported_rows(path: Path, sheet_name: str) -> list[tuple[str, str]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
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
        if carrier not in {"DGF", "DSV", "MAERSK"}:
            continue
        tracking_number = str(row[bill_idx] or "").strip()
        if tracking_number:
            rows.append((carrier, tracking_number))
    return rows


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
    if carrier == "DSV":
        shipment = client.track(tracking_number)
        return _dsv_record(tracking_number, shipment)
    raise ValueError(f"Unsupported carrier: {carrier}")


def _dsv_record(tracking_number: str, shipment) -> TrackingRecord:
    data = shipment.to_dict()
    if not data.get("found"):
        return TrackingRecord(carrier="DSV", tracking_number=tracking_number, found=False, raw={"payload": data})
    from datetime import datetime

    arrival = _parse_datetime(data.get("to_date"))
    actual_delivery = _parse_datetime(data.get("actual_delivery"))
    actual_pickup = _parse_datetime(data.get("actual_pickup"))
    arrival_date = actual_delivery or arrival
    return TrackingRecord(
        carrier="DSV",
        tracking_number=tracking_number,
        found=True,
        status=data.get("status"),
        eta_arrival=arrival,
        actual_arrival=actual_delivery,
        arrival_date=arrival_date,
        arrival_date_type="ACTUAL" if actual_delivery else "ESTIMATED" if arrival else None,
        actual_pickup=actual_pickup,
        actual_delivery=actual_delivery,
        house_bill=data.get("house_bill_number"),
        raw={"payload": data},
    )


def _parse_datetime(value):
    if not isinstance(value, str) or not value:
        return None
    from datetime import datetime

    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None


def _remark(record: TrackingRecord) -> str | None:
    if not record.found:
        return "NOT_FOUND: API returned no shipment"
    if not record.arrival_date:
        return "NO_ARRIVAL_DATE: shipment found but no ETA/ATA"
    return None


def _progress_line(index: int, total: int, record: TrackingRecord, remark: str | None) -> str:
    if remark:
        result = remark
    elif record.arrival_date:
        result = f"{record.arrival_date_type} {record.arrival_date:%Y-%m-%d %H:%M}"
    else:
        result = "OK"
    return f"[{record.carrier}] {index}/{total} {record.tracking_number} -> {result}"


if __name__ == "__main__":
    main()
