from __future__ import annotations

from dataclasses import dataclass
import json
from typing import Any
from urllib import parse, request


MYDSV_BASE_URL = "https://mydsv.com"


@dataclass(frozen=True)
class DsvShipment:
    query: str
    search_hit_count: int
    summary: dict[str, Any] | None
    detail: dict[str, Any] | None

    @property
    def found(self) -> bool:
        return self.summary is not None

    def to_dict(self) -> dict[str, Any]:
        summary = self.summary or {}
        detail = self.detail or {}
        return {
            "query": self.query,
            "found": self.found,
            "search_hit_count": self.search_hit_count,
            "main_tms_shipment_id": summary.get("mainTmsShipmentId"),
            "house_bill_number": detail.get("houseBillNumber") or summary.get("houseBillNumber"),
            "random_identifier": summary.get("randomIdentifier"),
            "status": detail.get("status") or summary.get("status"),
            "mode_of_transport": detail.get("modeOfTransport") or summary.get("modeOfTransport"),
            "from_date": detail.get("fromDate") or summary.get("fromDate"),
            "to_date": detail.get("toDate") or summary.get("toDate"),
            "actual_pickup": detail.get("actualPickup") or summary.get("actualPickup"),
            "actual_delivery": detail.get("actualDelivery") or summary.get("actualDelivery"),
            "milestones": detail.get("milestones") or [],
            "summary": summary,
        }


class DsvPublicClient:
    """Small client for the myDSV public tracking endpoints used by the web UI."""

    def __init__(self, base_url: str = MYDSV_BASE_URL, timeout: int = 30) -> None:
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout

    def track(self, house_bill_or_shipment_id: str) -> DsvShipment:
        query = house_bill_or_shipment_id.strip()
        hits = self.search(query)
        if not hits:
            return DsvShipment(query=query, search_hit_count=0, summary=None, detail=None)

        summary = hits[0]
        random_identifier = summary.get("randomIdentifier")
        detail = self.get_detail(random_identifier) if random_identifier else None
        return DsvShipment(
            query=query,
            search_hit_count=len(hits),
            summary=summary,
            detail=detail,
        )

    def search(self, query: str) -> list[dict[str, Any]]:
        params = parse.urlencode({"q": query})
        payload = self._get_json(f"/app/search/publicShipmentList?{params}")
        data = payload.get("data") if isinstance(payload, dict) else None
        return data if isinstance(data, list) else []

    def get_detail(self, random_identifier: str) -> dict[str, Any]:
        path = f"/app/search/shipment/{parse.quote(random_identifier)}?detailedMilestones=true"
        return self._get_json(path)

    def _get_json(self, path: str) -> dict[str, Any]:
        url = f"{self.base_url}{path}"
        req = request.Request(
            url,
            headers={
                "Accept": "application/json, text/plain, */*",
                "User-Agent": "Mozilla/5.0",
                "Referer": f"{self.base_url}/new/tracking/track-shipment",
            },
        )
        with request.urlopen(req, timeout=self.timeout) as resp:
            raw = resp.read().decode("utf-8")
        loaded = json.loads(raw)
        if not isinstance(loaded, dict):
            raise ValueError(f"Expected object response from {url}")
        return loaded
