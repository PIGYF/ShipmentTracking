from __future__ import annotations

from dataclasses import dataclass
import json
import os
from typing import Any
from urllib import parse, request


DEFAULT_TOKEN_URL = "https://api.maersk.com/customer-identity/oauth/v2/access_token"
DEFAULT_API_BASE_URL = "https://api.maersk.com"
DEFAULT_EVENTS_PATH = "/track-and-trace/events"


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

    def track(self, reference: str) -> MaerskTrackingResult:
        last_result: MaerskTrackingResult | None = None
        for query_type in ("transportDocumentReference", "carrierBookingReference", "equipmentReference"):
            result = self.track_by(reference, query_type)
            if result.found:
                return result
            last_result = result
        assert last_result is not None
        return last_result

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
            headers={"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"},
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
