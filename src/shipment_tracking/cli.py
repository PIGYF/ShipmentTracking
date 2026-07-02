from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from .dgf import DgfClient
from .dsv_public import DsvPublicClient
from .env import load_dotenv
from .excel_source import read_shipments_from_excel
from .excel_writer import update_tracking_workbook
from .maersk import MaerskClient
from .models import TrackingRecord


def main() -> None:
    load_dotenv()

    parser = argparse.ArgumentParser(description="Fetch shipment tracking data.")
    parser.add_argument("--excel", help="Excel file containing shipment number and forwarder columns.")
    parser.add_argument("--sheet", help="Optional worksheet name.")
    parser.add_argument(
        "--carrier",
        choices=["dsv", "maersk", "dgf", "dhl", "all"],
        default="dsv",
        help="Carrier rows to read from Excel. Default: dsv.",
    )
    parser.add_argument("--dgf", action="append", help="DGF tracking number. Can be repeated.")
    parser.add_argument("--dhl", action="append", help="Deprecated alias for --dgf.")
    parser.add_argument("--dsv", action="append", help="DSV House Bill / shipment id. Can be repeated.")
    parser.add_argument("--maersk", action="append", help="Maersk B/L, booking, or container reference. Can be repeated.")
    parser.add_argument("--output", default="exports/results.json", help="JSON output path.")
    parser.add_argument("--update-excel", help="Write normalized tracking results into an Excel copy.")
    args = parser.parse_args()

    results: list[dict[str, Any]] = []
    records: list[TrackingRecord] = []
    dgf_queries, dsv_queries, maersk_queries = _collect_queries(
        args.excel,
        args.sheet,
        args.carrier,
        args.dgf,
        args.dhl,
        args.dsv,
        args.maersk,
    )

    if dgf_queries:
        client = DgfClient()
        for query in dgf_queries:
            try:
                shipment = client.track(query)
            except RuntimeError as exc:
                raise SystemExit(f"DGF configuration error: {exc}") from exc
            item = shipment.to_dict()
            record = shipment.to_record()
            item["normalized"] = record.to_dict()
            item["carrier"] = "DGF"
            results.append(item)
            records.append(record)
            print(f"DGF {query}: {item.get('status') or 'NOT_FOUND'}")

    if dsv_queries:
        client = DsvPublicClient()
        for query in dsv_queries:
            shipment = client.track(query)
            item = shipment.to_dict()
            item["carrier"] = "DSV"
            results.append(item)
            print(f"DSV {query}: {item.get('status') or 'NOT_FOUND'}")

    if maersk_queries:
        client = MaerskClient()
        for query in maersk_queries:
            try:
                shipment = client.track(query)
            except RuntimeError as exc:
                raise SystemExit(f"Maersk configuration error: {exc}") from exc
            item = shipment.to_dict()
            item["carrier"] = "MAERSK"
            results.append(item)
            print(f"MAERSK {query}: {'FOUND' if item.get('found') else 'NOT_FOUND'} ({item.get('query_type')})")

    if not results:
        raise SystemExit("No shipments to query. Provide --excel, --dgf, --dsv, or --maersk.")

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {output_path}")

    if args.update_excel:
        if not args.excel:
            raise SystemExit("--update-excel requires --excel.")
        updated = update_tracking_workbook(args.excel, records, args.update_excel, args.sheet)
        print(f"Updated {updated} Excel rows: {args.update_excel}")


def _collect_queries(
    excel_path: str | None,
    sheet_name: str | None,
    carrier: str,
    dgf_values: list[str] | None,
    dhl_values: list[str] | None,
    dsv_values: list[str] | None,
    maersk_values: list[str] | None,
) -> tuple[list[str], list[str], list[str]]:
    dgf_queries = list(dgf_values or []) + list(dhl_values or [])
    dsv_queries = list(dsv_values or [])
    maersk_queries = list(maersk_values or [])

    if excel_path:
        for shipment in read_shipments_from_excel(excel_path, sheet_name):
            forwarder = shipment.forwarder.upper()
            if carrier in ("dgf", "dhl", "all") and forwarder == "DGF":
                dgf_queries.append(shipment.tracking_number)
            if carrier in ("dsv", "all") and forwarder == "DSV":
                dsv_queries.append(shipment.tracking_number)
            if carrier in ("maersk", "all") and forwarder == "MAERSK":
                maersk_queries.append(shipment.tracking_number)

    return _dedupe(dgf_queries), _dedupe(dsv_queries), _dedupe(maersk_queries)


def _dedupe(values: list[str]) -> list[str]:
    deduped: list[str] = []
    for value in values:
        clean = value.strip()
        if clean and clean not in deduped:
            deduped.append(clean)
    return deduped


if __name__ == "__main__":
    main()
