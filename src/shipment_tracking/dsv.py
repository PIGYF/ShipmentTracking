from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
import os
from typing import Any
from urllib import parse, request
from urllib.error import HTTPError

from .destination import destination_aliases, matches_destination
from .models import TrackingRecord
from .time_utils import to_china_naive


DEFAULT_TOKEN_URL = "https://api.dsv.com/my/oauth/v1/token"
DEFAULT_API_BASE_URL = "https://api.dsv.com/my/tracking/v2"


@dataclass(frozen=True)
class DsvTrackingResult:
    query: str
    status_code: int
    payload: Any

    @property
    def found(self) -> bool:
        return self.status_code == 200 and bool(_first_shipment(self.payload))

    def to_record(self) -> TrackingRecord:
        shipment = _first_shipment(self.payload)
        if not shipment:
            return TrackingRecord(
                carrier="DSV",
                tracking_number=self.query,
                found=False,
                raw={"payload": self.payload} if isinstance(self.payload, dict) else {},
            )

        events = shipment.get("events") if isinstance(shipment.get("events"), list) else []
        aliases = destination_aliases()
        booking_event = _earliest_event(events, "BOOKING")
        actual_pickup_event = _latest_event(events, "PCF") or _latest_event(events, "PUP")
        departure_event = _earliest_event(events, "DEP") or _earliest_event(events, "ETD")
        actual_arrival_event = _latest_destination_event(events, "ARV", aliases)
        eta_arrival_event = _latest_destination_event(events, "ETA", aliases)

        call_for_pickup_date = to_china_naive(_parse_datetime(booking_event.get("eventDate")) if booking_event else None)
        actual_pickup = to_china_naive(_parse_datetime(actual_pickup_event.get("eventDate")) if actual_pickup_event else None)
        departure_date = to_china_naive(_parse_datetime(departure_event.get("eventDate")) if departure_event else None)
        actual_arrival = to_china_naive(_parse_datetime(actual_arrival_event.get("eventDate")) if actual_arrival_event else None)
        eta_arrival = to_china_naive(_parse_datetime(eta_arrival_event.get("eventDate")) if eta_arrival_event else None)
        arrival_date = actual_arrival or eta_arrival
        arrival_date_type = "ACTUAL" if actual_arrival else "ESTIMATED" if eta_arrival else None
        arrival_event = actual_arrival_event or eta_arrival_event or {}

        return TrackingRecord(
            carrier="DSV",
            tracking_number=self.query,
            found=True,
            status=_string(shipment.get("status")),
            status_description=_string(arrival_event.get("eventDescription") or shipment.get("status")),
            eta_arrival=eta_arrival,
            actual_arrival=actual_arrival,
            arrival_date=arrival_date,
            arrival_date_type=arrival_date_type,
            call_for_pickup_date=call_for_pickup_date,
            departure_date=departure_date,
            actual_pickup=actual_pickup,
            origin=_event_location(departure_event or {}),
            destination=_event_location(arrival_event),
            house_bill=self.query,
            master_bill=_string(shipment.get("shipmentId")),
            container_numbers=_container_numbers(events),
            vessel_name=_string(arrival_event.get("vesselName")),
            voyage_number=_string(arrival_event.get("voyageFlightNo")),
            raw={"payload": self.payload} if isinstance(self.payload, dict) else {},
        )


class DsvClient:
    """DSV production tracking API client using myDSV OAuth credentials."""

    def __init__(
        self,
        access_token_subscription_key: str | None = None,
        tracking_subscription_key: str | None = None,
        client_id: str | None = None,
        client_secret: str | None = None,
        token_url: str | None = None,
        api_base_url: str | None = None,
        timeout: int = 30,
    ) -> None:
        self.access_token_subscription_key = access_token_subscription_key or os.getenv("DSV_ACCESS_TOKEN_SUBSCRIPTION_KEY", "")
        self.tracking_subscription_key = tracking_subscription_key or os.getenv("DSV_TRACKING_SUBSCRIPTION_KEY", "")
        self.client_id = client_id or os.getenv("DSV_CLIENT_ID", "")
        self.client_secret = client_secret or os.getenv("DSV_CLIENT_SECRET", "")
        self.token_url = token_url or os.getenv("DSV_TOKEN_URL", DEFAULT_TOKEN_URL)
        self.api_base_url = (api_base_url or os.getenv("DSV_API_BASE_URL", DEFAULT_API_BASE_URL)).rstrip("/")
        self.timeout = timeout
        self._access_token: str | None = None

        missing = [
            name
            for name, value in (
                ("DSV_ACCESS_TOKEN_SUBSCRIPTION_KEY", self.access_token_subscription_key),
                ("DSV_TRACKING_SUBSCRIPTION_KEY", self.tracking_subscription_key),
                ("DSV_CLIENT_ID", self.client_id),
                ("DSV_CLIENT_SECRET", self.client_secret),
            )
            if not value.strip()
        ]
        if missing:
            raise RuntimeError(f"Missing DSV configuration: {', '.join(missing)}")

    def track(self, reference: str) -> DsvTrackingResult:
        query = reference.strip()
        token = self._get_access_token()
        url = f"{self.api_base_url}/shipments/reference/{parse.quote(query)}"
        status_code, payload = self._get_json(url, token)
        return DsvTrackingResult(query=query, status_code=status_code, payload=payload)

    def _get_access_token(self) -> str:
        if self._access_token:
            return self._access_token

        form = parse.urlencode(
            {
                "grant_type": "client_credentials",
                "client_id": self.client_id.strip(),
                "client_secret": self.client_secret.strip(),
            }
        ).encode("utf-8")
        req = request.Request(
            self.token_url,
            data=form,
            headers={
                "Accept": "application/json",
                "Content-Type": "application/x-www-form-urlencoded",
                "DSV-Subscription-Key": self.access_token_subscription_key.strip(),
            },
            method="POST",
        )
        try:
            with request.urlopen(req, timeout=self.timeout) as resp:
                payload = _read_json(resp.read())
        except HTTPError as exc:
            raise RuntimeError(f"DSV token API rejected credentials or permissions (HTTP {exc.code})") from exc

        token = payload.get("access_token") if isinstance(payload, dict) else None
        if not token:
            raise RuntimeError("DSV token response did not include access_token.")
        self._access_token = token
        return token

    def _get_json(self, url: str, token: str) -> tuple[int, Any]:
        req = request.Request(
            url,
            headers={
                "Accept": "application/json",
                "Authorization": f"Bearer {token}",
                "DSV-Subscription-Key": self.tracking_subscription_key.strip(),
            },
        )
        try:
            with request.urlopen(req, timeout=self.timeout) as resp:
                return resp.status, _read_json(resp.read())
        except HTTPError as exc:
            if exc.code in {401, 403}:
                raise RuntimeError(f"DSV tracking API rejected credentials or permissions (HTTP {exc.code})") from exc
            return exc.code, _read_json(exc.read())


def _read_json(raw: bytes) -> Any:
    text = raw.decode("utf-8", errors="replace")
    if not text:
        return {}
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return {"raw": text}


def _first_shipment(payload: Any) -> dict[str, Any] | None:
    if isinstance(payload, list):
        for item in payload:
            if isinstance(item, dict):
                return item
    if isinstance(payload, dict) and payload.get("shipmentId"):
        return payload
    return None


def _latest_event(events: list[Any], code: str) -> dict[str, Any] | None:
    matches = [
        event
        for event in events
        if isinstance(event, dict)
        and str(event.get("eventCode") or "").upper() == code
        and not event.get("cancelled")
    ]
    if not matches:
        return None
    return max(matches, key=lambda event: _parse_datetime(event.get("eventDate")) or datetime.min)


def _earliest_event(events: list[Any], code: str) -> dict[str, Any] | None:
    matches = [
        event
        for event in events
        if isinstance(event, dict)
        and str(event.get("eventCode") or "").upper() == code
        and not event.get("cancelled")
    ]
    if not matches:
        return None
    return min(matches, key=lambda event: _parse_datetime(event.get("eventDate")) or datetime.max)


def _latest_destination_event(events: list[Any], code: str, aliases: tuple[str, ...]) -> dict[str, Any] | None:
    matches = [
        event
        for event in events
        if isinstance(event, dict)
        and str(event.get("eventCode") or "").upper() == code
        and not event.get("cancelled")
        and (matches_destination(event.get("location"), aliases) or matches_destination(event, aliases))
    ]
    if not matches:
        return None
    return max(matches, key=lambda event: _parse_datetime(event.get("eventDate")) or datetime.min)


def _parse_datetime(value: Any) -> datetime | None:
    if not isinstance(value, str) or not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None


def _event_location(event: dict[str, Any]) -> str | None:
    location = event.get("location") if isinstance(event.get("location"), dict) else {}
    parts = [_string(location.get("place")), _string(location.get("countryCode"))]
    return ", ".join(part for part in parts if part) or None


def _container_numbers(events: list[Any]) -> list[str]:
    numbers: list[str] = []
    for event in events:
        if not isinstance(event, dict) or event.get("eventType") != "CONTAINER":
            continue
        value = _string(event.get("eventTypeValue"))
        if value and value not in numbers:
            numbers.append(value)
    return numbers


def _string(value: Any) -> str | None:
    if value is None:
        return None
    return str(value)
