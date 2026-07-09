from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
import os
from typing import Any
from urllib import parse, request

from .destination import destination_aliases, matches_destination
from .models import TrackingRecord
from .time_utils import to_china_naive


DEFAULT_API_BASE_URL = "https://api-eu.dhl.com/track"


@dataclass(frozen=True)
class DgfTrackingResult:
    query: str
    service: str | None
    status_code: int
    payload: Any

    @property
    def found(self) -> bool:
        if self.status_code != 200:
            return False
        if isinstance(self.payload, dict):
            shipments = self.payload.get("shipments")
            return isinstance(shipments, list) and bool(shipments)
        return False

    def to_dict(self) -> dict[str, Any]:
        shipment = None
        if isinstance(self.payload, dict):
            shipments = self.payload.get("shipments")
            if isinstance(shipments, list) and shipments:
                shipment = shipments[0]

        return {
            "query": self.query,
            "service": self.service,
            "found": self.found,
            "status_code": self.status_code,
            "status": _first(shipment, "status", "statusCode"),
            "status_description": _first(shipment, "status", "description"),
            "estimated_delivery": _first(shipment, "estimatedTimeOfDelivery"),
            "origin": _first(shipment, "origin", "address", "addressLocality"),
            "destination": _first(shipment, "destination", "address", "addressLocality"),
            "events": _first(shipment, "events") or [],
            "payload": self.payload,
        }

    def to_record(self) -> TrackingRecord:
        shipment = _first_shipment(self.payload)
        if not shipment:
            return TrackingRecord(
                carrier="DGF",
                tracking_number=self.query,
                found=False,
                raw={"payload": self.payload} if isinstance(self.payload, dict) else {},
            )

        details = shipment.get("details") if isinstance(shipment.get("details"), dict) else {}
        aliases = destination_aliases()
        route = _first(details, "dgf:routes") or []
        route_item = _destination_route(route, aliases)
        events = shipment.get("events") if isinstance(shipment.get("events"), list) else []

        actual_arrival = _actual_destination_arrival(events, aliases)
        eta_arrival = _estimated_destination_arrival(events, route_item, aliases)
        actual_arrival = to_china_naive(actual_arrival)
        eta_arrival = to_china_naive(eta_arrival)
        call_for_pickup_date = to_china_naive(_find_event_timestamp(events, "estimated pick up date"))
        actual_pickup = to_china_naive(_find_event_timestamp(events, "actual pickup date"))
        departure_date = to_china_naive(_departure_timestamp(events, route_item))
        arrival_date = actual_arrival or eta_arrival
        arrival_date_type = "ACTUAL" if actual_arrival else "ESTIMATED" if eta_arrival else None

        references = details.get("references") if isinstance(details.get("references"), list) else []
        master_bills = _reference_numbers(references, "masterbill")
        containers = _reference_numbers(references, "container-number")
        house_bills = _reference_numbers(references, "housebill")

        return TrackingRecord(
            carrier="DGF",
            tracking_number=self.query,
            found=True,
            status=_first(shipment, "status", "statusCode"),
            status_description=_first(shipment, "status", "description"),
            eta_arrival=eta_arrival,
            actual_arrival=actual_arrival,
            arrival_date=arrival_date,
            arrival_date_type=arrival_date_type,
            call_for_pickup_date=call_for_pickup_date,
            departure_date=departure_date,
            actual_pickup=actual_pickup,
            origin=_location_name(shipment.get("origin")),
            destination=_route_destination_name(route_item) or _location_name(shipment.get("destination")),
            master_bill=master_bills[0] if master_bills else None,
            house_bill=house_bills[0] if house_bills else self.query,
            container_numbers=containers,
            vessel_name=_first(route_item, "dgf:vesselName"),
            voyage_number=_first(route_item, "dgf:voyageFlightNumber"),
            raw={"payload": self.payload} if isinstance(self.payload, dict) else {},
        )


class DgfClient:
    """DGF tracking client using DHL Shipment Tracking - Unified API."""

    def __init__(
        self,
        api_key: str | None = None,
        api_base_url: str | None = None,
        timeout: int = 30,
    ) -> None:
        self.api_key = api_key or os.getenv("DGF_API_KEY", "")
        self.api_base_url = (api_base_url or os.getenv("DGF_API_BASE_URL", DEFAULT_API_BASE_URL)).rstrip("/")
        self.timeout = timeout

    def track(self, tracking_number: str, service: str | None = None) -> DgfTrackingResult:
        if not self.api_key:
            raise RuntimeError("Missing DGF_API_KEY.")

        params = {"trackingNumber": tracking_number.strip()}
        if service:
            params["service"] = service
        url = f"{self.api_base_url}/shipments?{parse.urlencode(params)}"
        status_code, payload = self._get_json(url)
        return DgfTrackingResult(
            query=tracking_number.strip(),
            service=service,
            status_code=status_code,
            payload=payload,
        )

    def _get_json(self, url: str) -> tuple[int, Any]:
        req = request.Request(
            url,
            headers={
                "Accept": "application/json",
                "DHL-API-Key": self.api_key,
            },
        )
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


def _first(value: Any, *keys: str) -> Any:
    current = value
    for key in keys:
        if not isinstance(current, dict):
            return None
        current = current.get(key)
    return current


def _first_shipment(payload: Any) -> dict[str, Any] | None:
    if not isinstance(payload, dict):
        return None
    shipments = payload.get("shipments")
    if not isinstance(shipments, list) or not shipments or not isinstance(shipments[0], dict):
        return None
    return shipments[0]


def _parse_datetime(value: Any) -> datetime | None:
    if not isinstance(value, str) or not value:
        return None
    clean = value.replace("Z", "+00:00")
    try:
        return datetime.fromisoformat(clean)
    except ValueError:
        return None


def _actual_destination_arrival(events: list[Any], aliases: tuple[str, ...]) -> datetime | None:
    return (
        _find_destination_event_timestamp(events, "actual vessel arrival", aliases)
        or _find_destination_event_timestamp(events, "actual arrival at destination", aliases)
    )


def _estimated_destination_arrival(events: list[Any], route_item: dict[str, Any], aliases: tuple[str, ...]) -> datetime | None:
    return (
        _find_destination_event_timestamp(events, "estimated vessel arrival (last updated)", aliases)
        or _find_destination_event_timestamp(events, "estimated vessel arrival", aliases)
        or (_parse_datetime(_first(route_item, "dgf:estimatedArrivalDate")) if matches_destination(route_item, aliases) else None)
    )


def _find_destination_event_timestamp(events: list[Any], description_fragment: str, aliases: tuple[str, ...]) -> datetime | None:
    fragment = description_fragment.lower()
    matches: list[datetime] = []
    for event in events:
        if not isinstance(event, dict):
            continue
        description = str(event.get("description") or "").lower()
        if fragment not in description:
            continue
        parsed = _parse_datetime(event.get("timestamp"))
        if not parsed:
            continue
        if matches_destination(event.get("location"), aliases) or matches_destination(event, aliases):
            matches.append(parsed)
    if matches:
        return max(matches)
    return None


def _find_event_timestamp(events: list[Any], description_fragment: str) -> datetime | None:
    fragment = description_fragment.lower()
    matches: list[datetime] = []
    for event in events:
        if not isinstance(event, dict):
            continue
        description = str(event.get("description") or "").lower()
        if fragment in description:
            parsed = _parse_datetime(event.get("timestamp"))
            if parsed:
                matches.append(parsed)
    if not matches:
        return None
    return max(matches)


def _departure_timestamp(events: list[Any], route_item: dict[str, Any]) -> datetime | None:
    return (
        _find_first_event_timestamp(events, "actual vessel departure")
        or _find_first_event_timestamp(events, "estimated vessel departure (last updated)")
        or _find_first_event_timestamp(events, "estimated vessel departure")
        or _parse_datetime(_first(route_item, "dgf:estimatedDepartureDate"))
    )


def _find_first_event_timestamp(events: list[Any], description_fragment: str) -> datetime | None:
    fragment = description_fragment.lower()
    matches: list[datetime] = []
    for event in events:
        if not isinstance(event, dict):
            continue
        description = str(event.get("description") or "").lower()
        if fragment in description:
            parsed = _parse_datetime(event.get("timestamp"))
            if parsed:
                matches.append(parsed)
    if not matches:
        return None
    return min(matches)


def _reference_numbers(references: list[Any], reference_type: str) -> list[str]:
    numbers: list[str] = []
    for reference in references:
        if not isinstance(reference, dict):
            continue
        if reference.get("type") != reference_type:
            continue
        number = str(reference.get("number") or "").strip()
        if number and number not in numbers:
            numbers.append(number)
    return numbers


def _location_name(value: Any) -> str | None:
    if not isinstance(value, dict):
        return None
    locality = _first(value, "address", "addressLocality")
    country = _first(value, "address", "countryCode")
    if locality and country:
        return f"{locality}, {country}"
    return locality or country


def _destination_route(routes: Any, aliases: tuple[str, ...]) -> dict[str, Any]:
    if not isinstance(routes, list):
        return {}
    route_items = [item for item in routes if isinstance(item, dict)]
    for route_item in route_items:
        destination = _first_existing(
            route_item,
            "destination",
            "dgf:destination",
            "to",
            "dgf:to",
            "arrival",
            "dgf:arrival",
        )
        if matches_destination(destination, aliases):
            return route_item
    for route_item in route_items:
        if matches_destination(route_item, aliases):
            return route_item
    return route_items[0] if route_items else {}


def _route_destination_name(route_item: dict[str, Any]) -> str | None:
    destination = _first_existing(
        route_item,
        "destination",
        "dgf:destination",
        "to",
        "dgf:to",
        "arrival",
        "dgf:arrival",
    )
    return _location_name(destination) or _string_from_location(destination)


def _first_existing(value: dict[str, Any], *keys: str) -> Any:
    for key in keys:
        if key in value and value[key]:
            return value[key]
    return None


def _string_from_location(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, str):
        return value
    if isinstance(value, dict):
        for key in (
            "addressLocality",
            "city",
            "cityName",
            "locationName",
            "place",
            "airportCode",
            "iataCode",
            "unLocationCode",
            "UNLocationCode",
            "code",
            "locationCode",
            "dgf:arrivalAirportCode",
            "dgf:destinationAirportCode",
            "dgf:arrivalCity",
            "dgf:destinationCity",
            "dgf:arrivalLocationCode",
            "dgf:destinationLocationCode",
        ):
            if value.get(key):
                return str(value[key])
    return None
