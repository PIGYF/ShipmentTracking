from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from .dsv_public import DsvPublicClient
from .excel_source import read_shipments_from_excel


def main() -> None:
    parser = argparse.ArgumentParser(description="Fetch shipment tracking data.")
    parser.add_argument("--excel", help="Excel file containing 运单 and 货代 columns.")
    parser.add_argument("--dsv", action="append", help="DSV House Bill / shipment id to query. Can be repeated.")
    parser.add_argument("--output", default="exports/results.json", help="JSON output path.")
    args = parser.parse_args()

    queries = _collect_queries(args.excel, args.dsv)
    client = DsvPublicClient()
    results: list[dict[str, Any]] = []

    for query in queries:
        shipment = client.track(query)
        results.append(shipment.to_dict())
        status = results[-1].get("status") or "NOT_FOUND"
        print(f"{query}: {status}")

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {output_path}")


def _collect_queries(excel_path: str | None, dsv_values: list[str] | None) -> list[str]:
    queries: list[str] = []
    if excel_path:
        for shipment in read_shipments_from_excel(excel_path):
            if shipment.forwarder.upper() == "DSV":
                queries.append(shipment.tracking_number)
    if dsv_values:
        queries.extend(dsv_values)

    deduped: list[str] = []
    for query in queries:
        if query and query not in deduped:
            deduped.append(query)
    if not deduped:
        raise SystemExit("No DSV shipments to query. Provide --excel or --dsv.")
    return deduped


if __name__ == "__main__":
    main()
