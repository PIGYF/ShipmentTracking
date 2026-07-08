from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
import os
import re
from typing import Any
from urllib import parse, request

from .models import TrackingRecord
from .time_utils import to_china_naive


DEFAULT_TOKEN_URL = "https://api.maersk.com/customer-identity/oauth/v2/access_token"
DEFAULT_API_BASE_URL = "https://api.maersk.com/track-and-trace-private"
DEFAULT_EVENTS_PATH = "/events"
QUERY_TYPES = ("transportDocumentReference", "carrierBookingReference", "equipmentReference")
CONTAINER_RE = re.compile(r"^[A-Z]{4}\d{7}$")


@dataclass(frozen=True)
class MaerskTrackingResult:
    query: str
    query_type: str
    status_code: int
    payload: Any

    @property
    def found(self) -> bool:
        if self.status_code != 200:
            return False
        if isinstance(self.payload, list):
            return bool(self.payload)
        if isinstance(self.payload, dict):
            for key in ("events", "data", "transportEvents", "shipmentEvents", "equipmentEvents"):
                value = self.payload.get(key)
                if isinstance(value, list) and value:
                    return True
            return bool(self.payload)
        return False

    def to_dict(self) -> dict[str, Any]:
        return {
            "query": self.query,
            "query_type": self.query_type,
            "found": self.found,
            "status_code": self.status_code,
            "payload": self.payload,
        }

    def to_record(self) -> TrackingRecord:
        events = _events(self.payload)
        if not self.found:
            return TrackingRecord(
                carrier="MAERSK",
                tracking_number=self.query,
                found=False,
                raw={"payload": self.payload} if isinstance(self.payload, dict) else {},
            )

        final_sequence = _final_arrival_sequence(events)
        eta_arrival_event = _transport_arrival_at_sequence(events, "EST", final_sequence)
        actual_arrival_event = _transport_arrival_at_sequence(events, "ACT", final_sequence)
        actual_discharge_event = _equipment_discharge_at_sequence(events, final_sequence)
        actual_pickup_event = _first_equipment_event(events, "GTIN", "ACT")
        departure_event = _origin_departure_event(events)

        eta_arrival = to_china_naive(_parse_datetime(eta_arrival_event.get("eventDateTime")) if eta_arrival_event else None)
        actual_arrival = to_china_naive(_parse_datetime(actual_arrival_event.get("eventDateTime")) if actual_arrival_event else None)
        fallback_discharge = to_china_naive(_parse_datetime(actual_discharge_event.get("eventDateTime")) if actual_discharge_event else None)
        actual_pickup = to_china_naive(_parse_datetime(actual_pickup_event.get("eventDateTime")) if actual_pickup_event else None)
        departure_date = to_china_naive(_parse_datetime(departure_event.get("eventDateTime")) if departure_event else None)
        arrival_date = actual_arrival or eta_arrival or fallback_discharge
        arrival_date_type = "ACTUAL" if actual_arrival or fallback_discharge else "ESTIMATED" if eta_arrival else None

        final_event = actual_arrival_event or eta_arrival_event or actual_discharge_event or {}
        transport_call = final_event.get("transportCall") if isinstance(final_event.get("transportCall"), dict) else {}
        vessel = transport_call.get("vessel") if isinstance(transport_call.get("vessel"), dict) else {}

        return TrackingRecord(
            carrier="MAERSK",
            tracking_number=self.query,
            found=True,
            status=_latest_classifier(events),
            eta_arrival=eta_arrival,
            actual_arrival=actual_arrival or fallback_discharge,
            arrival_date=arrival_date,
            arrival_date_type=arrival_date_type,
            departure_date=departure_date,
            actual_pickup=actual_pickup,
            destination=_transport_location(transport_call),
            master_bill=self.query if self.query_type == "transportDocumentReference" else None,
            container_numbers=_reference_values(events, "EQ"),
            vessel_name=vessel.get("vesselName"),
            voyage_number=transport_call.get("exportVoyageNumber") or transport_call.get("carrierVoyageNumber"),
            raw={"payload": self.payload} if isinstance(self.payload, dict) else {},
        )


class MaerskClient:
    """Template client for Maersk Track & Trace Events API."""

    def __init__(
        self,
        consumer_key: str | None = None,
        consumer_secret: str | None = None,
        customer_key: str | None = None,
        token_url: str | None = None,
        api_base_url: str | None = None,
        events_path: str | None = None,
        timeout: int = 30,
    ) -> None:
        self.consumer_key = consumer_key or os.getenv("MAERSK_CONSUMER_KEY", "")
        self.consumer_secret = consumer_secret or os.getenv("MAERSK_CONSUMER_SECRET", "")
        self.customer_key = customer_key or os.getenv("MAERSK_CUSTOMER_KEY", "")
        self.token_url = token_url or os.getenv("MAERSK_TOKEN_URL", DEFAULT_TOKEN_URL)
        self.api_base_url = (api_base_url or os.getenv("MAERSK_API_BASE_URL", DEFAULT_API_BASE_URL)).rstrip("/")
        self.events_path = events_path or os.getenv("MAERSK_EVENTS_PATH", DEFAULT_EVENTS_PATH)
        self.timeout = timeout
        self._access_token: str | None = None
        self._successful_query_type_by_kind: dict[str, str] = {}

    def track(self, reference: str) -> MaerskTrackingResult:
        last_result: MaerskTrackingResult | None = None
        reference = reference.strip()
        kind = _reference_kind(reference)
        for query_type in self._query_order(reference):
            result = self.track_by(reference, query_type)
            if result.found:
                self._successful_query_type_by_kind[kind] = query_type
                return result
            last_result = result
        assert last_result is not None
        return last_result

    def _query_order(self, reference: str) -> tuple[str, ...]:
        kind = _reference_kind(reference)
        learned = self._successful_query_type_by_kind.get(kind)
        if learned:
            return _prefer(learned, QUERY_TYPES)
        if kind == "container":
            return ("equipmentReference", "transportDocumentReference", "carrierBookingReference")
        if kind == "numeric":
            return ("carrierBookingReference", "transportDocumentReference", "equipmentReference")
        return QUERY_TYPES

    def track_by(self, reference: str, query_type: str) -> MaerskTrackingResult:
        params = parse.urlencode({query_type: reference.strip()})
        url = f"{self.api_base_url}{self.events_path}?{params}"
        status_code, payload = self._get_json(url)
        return MaerskTrackingResult(
            query=reference.strip(),
            query_type=query_type,
            status_code=status_code,
            payload=payload,
        )

    def _get_access_token(self) -> str:
        if self._access_token:
            return self._access_token
        if not self.consumer_key or not self.consumer_secret:
            raise RuntimeError("Missing MAERSK_CONSUMER_KEY or MAERSK_CONSUMER_SECRET.")

        form = parse.urlencode(
            {
                "grant_type": "client_credentials",
                "client_id": self.consumer_key,
                "client_secret": self.consumer_secret,
            }
        ).encode("utf-8")
        req = request.Request(
            self.token_url,
            data=form,
            headers={
                "Content-Type": "application/x-www-form-urlencoded",
                "Accept": "application/json",
                "Consumer-Key": self.consumer_key,
            },
            method="POST",
        )
        with request.urlopen(req, timeout=self.timeout) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
        token = payload.get("access_token")
        if not token:
            raise RuntimeError("Maersk token response did not include access_token.")
        self._access_token = token
        return token

    def _get_json(self, url: str) -> tuple[int, Any]:
        token = self._get_access_token()
        headers = {
            "Accept": "application/json",
            "Authorization": f"Bearer {token}",
            "Consumer-Key": self.consumer_key,
        }
        if self.customer_key:
            headers["Customer-Key"] = self.customer_key

        req = request.Request(url, headers=headers)
        try:
            with request.urlopen(req, timeout=self.timeout) as resp:
                return resp.status, _read_json(resp.read())
        except Exception as exc:
            status = getattr(exc, "code", 0) or 0
            body = exc.read() if hasattr(exc, "read") else str(exc).encode("utf-8")
            return int(status), _read_json(body)


def _read_json(raw: bytes) -> Any:
    text = raw.decode("utf-8", errors="replace")
    if not text:
        return {}
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return {"raw": text}


def _events(payload: Any) -> list[dict[str, Any]]:
    if isinstance(payload, list):
        return [event for event in payload if isinstance(event, dict)]
    if not isinstance(payload, dict):
        return []
    found: list[dict[str, Any]] = []
    for key in ("events", "data", "transportEvents", "shipmentEvents", "equipmentEvents"):
        value = payload.get(key)
        if isinstance(value, list):
            found.extend(event for event in value if isinstance(event, dict))
    return found


def _final_arrival_sequence(events: list[dict[str, Any]]) -> int:
    sequences = [
        _sequence(event)
        for event in events
        if event.get("eventType") == "TRANSPORT" and event.get("transportEventTypeCode") == "ARRI"
    ]
    valid = [sequence for sequence in sequences if sequence >= 0]
    return max(valid) if valid else -1


def _transport_arrival_at_sequence(
    events: list[dict[str, Any]],
    classifier: str,
    sequence: int,
) -> dict[str, Any] | None:
    candidates = [
        event
        for event in events
        if event.get("eventType") == "TRANSPORT"
        and event.get("transportEventTypeCode") == "ARRI"
        and event.get("eventClassifierCode") == classifier
        and _sequence(event) == sequence
    ]
    return _final_event_by_sequence(candidates)


def _equipment_discharge_at_sequence(events: list[dict[str, Any]], sequence: int) -> dict[str, Any] | None:
    candidates = [
        event
        for event in events
        if event.get("eventType") == "EQUIPMENT"
        and event.get("equipmentEventTypeCode") == "DISC"
        and event.get("eventClassifierCode") == "ACT"
        and _sequence(event) == sequence
    ]
    return _final_event_by_sequence(candidates)


def _first_equipment_event(events: list[dict[str, Any]], event_type_code: str, classifier: str) -> dict[str, Any] | None:
    candidates = [
        event
        for event in events
        if event.get("eventType") == "EQUIPMENT"
        and event.get("equipmentEventTypeCode") == event_type_code
        and event.get("eventClassifierCode") == classifier
    ]
    return _first_event_by_datetime(candidates)


def _origin_departure_event(events: list[dict[str, Any]]) -> dict[str, Any] | None:
    actual = _origin_transport_departure(events, "ACT")
    if actual:
        return actual
    return _origin_transport_departure(events, "EST")


def _origin_transport_departure(events: list[dict[str, Any]], classifier: str) -> dict[str, Any] | None:
    candidates = [
        event
        for event in events
        if event.get("eventType") == "TRANSPORT"
        and event.get("transportEventTypeCode") == "DEPA"
        and event.get("eventClassifierCode") == classifier
        and _sequence(event) >= 2
    ]
    if not candidates:
        candidates = [
            event
            for event in events
            if event.get("eventType") == "TRANSPORT"
            and event.get("transportEventTypeCode") == "DEPA"
            and event.get("eventClassifierCode") == classifier
        ]
    return _first_event_by_datetime(candidates)


def _final_event_by_sequence(events: list[dict[str, Any]]) -> dict[str, Any] | None:
    if not events:
        return None
    return max(events, key=lambda event: (_sequence(event), _parse_datetime(event.get("eventDateTime")) or datetime.min))


def _first_event_by_datetime(events: list[dict[str, Any]]) -> dict[str, Any] | None:
    dated = [(event, _parse_datetime(event.get("eventDateTime"))) for event in events]
    dated = [(event, value) for event, value in dated if value]
    if not dated:
        return None
    event, _ = min(dated, key=lambda item: item[1])
    return event


def _sequence(event: dict[str, Any]) -> int:
    transport_call = event.get("transportCall")
    if not isinstance(transport_call, dict):
        return -1
    value = transport_call.get("transportCallSequenceNumber")
    try:
        return int(value)
    except (TypeError, ValueError):
        return -1


def _parse_datetime(value: Any) -> datetime | None:
    if not isinstance(value, str) or not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None


def _transport_location(transport_call: dict[str, Any]) -> str | None:
    location = transport_call.get("location") if isinstance(transport_call.get("location"), dict) else {}
    name = location.get("locationName")
    un_location = transport_call.get("UNLocationCode")
    if name and un_location:
        return f"{name}, {un_location}"
    return name or un_location


def _reference_values(events: list[dict[str, Any]], reference_type: str) -> list[str]:
    values: list[str] = []
    for event in events:
        references = event.get("references")
        if not isinstance(references, list):
            continue
        for reference in references:
            if not isinstance(reference, dict) or reference.get("referenceType") != reference_type:
                continue
            value = str(reference.get("referenceValue") or "").strip()
            if value and value not in values:
                values.append(value)
    return values


def _latest_classifier(events: list[dict[str, Any]]) -> str | None:
    dated = [(event, _parse_datetime(event.get("eventDateTime"))) for event in events]
    dated = [(event, value) for event, value in dated if value]
    if not dated:
        return None
    event, _ = max(dated, key=lambda item: item[1])
    parts = [event.get("eventType"), event.get("transportEventTypeCode") or event.get("equipmentEventTypeCode") or event.get("shipmentEventTypeCode"), event.get("eventClassifierCode")]
    return ":".join(str(part) for part in parts if part)


def _reference_kind(reference: str) -> str:
    clean = reference.strip().upper().replace(" ", "")
    if CONTAINER_RE.match(clean):
        return "container"
    if clean.isdigit():
        return "numeric"
    return "other"


def _prefer(first: str, values: tuple[str, ...]) -> tuple[str, ...]:
    return (first, *(value for value in values if value != first))
