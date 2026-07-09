from __future__ import annotations

import os
import re
from typing import Any


DEFAULT_DESTINATION_ALIASES = ("SHANGHAI", "上海", "PVG", "SHA", "CNSHA")


def destination_aliases() -> tuple[str, ...]:
    raw = os.getenv("SHIPMENT_DESTINATION_ALIASES") or os.getenv("DGF_DESTINATION_ALIASES")
    if not raw:
        return DEFAULT_DESTINATION_ALIASES
    aliases = tuple(alias.strip().upper() for alias in raw.split(",") if alias.strip())
    return aliases or DEFAULT_DESTINATION_ALIASES


def matches_destination(value: Any, aliases: tuple[str, ...]) -> bool:
    if value is None:
        return False
    if isinstance(value, dict):
        candidate_keys = {
            "addressLocality",
            "city",
            "cityName",
            "locationName",
            "place",
            "airportCode",
            "iataCode",
            "unLocationCode",
            "UNLocationCode",
            "UNLocationCode".lower(),
            "code",
            "locationCode",
            "countryCode",
            "dgf:arrivalAirportCode",
            "dgf:destinationAirportCode",
            "dgf:arrivalCity",
            "dgf:destinationCity",
            "dgf:arrivalLocationCode",
            "dgf:destinationLocationCode",
        }
        for key, child in value.items():
            if key in candidate_keys and string_matches_alias(child, aliases):
                return True
            if isinstance(child, (dict, list)) and matches_destination(child, aliases):
                return True
        return False
    if isinstance(value, list):
        return any(matches_destination(item, aliases) for item in value)
    return string_matches_alias(value, aliases)


def string_matches_alias(value: Any, aliases: tuple[str, ...]) -> bool:
    text = str(value or "").strip().upper()
    if not text:
        return False
    tokens = {token for token in re.split(r"[^A-Z0-9]+", text) if token}
    for alias in aliases:
        if len(alias) <= 3:
            if alias == text or alias in tokens:
                return True
            continue
        if alias in text:
            return True
    return False
