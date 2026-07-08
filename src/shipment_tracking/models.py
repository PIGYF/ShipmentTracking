from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import datetime
from typing import Any


@dataclass(frozen=True)
class TrackingRecord:
    carrier: str
    tracking_number: str
    found: bool
    status: str | None = None
    status_description: str | None = None
    eta_arrival: datetime | None = None
    actual_arrival: datetime | None = None
    arrival_date: datetime | None = None
    arrival_date_type: str | None = None
    call_for_pickup_date: datetime | None = None
    departure_date: datetime | None = None
    actual_pickup: datetime | None = None
    actual_delivery: datetime | None = None
    origin: str | None = None
    destination: str | None = None
    master_bill: str | None = None
    house_bill: str | None = None
    container_numbers: list[str] = field(default_factory=list)
    vessel_name: str | None = None
    voyage_number: str | None = None
    raw: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        for key, value in list(data.items()):
            if isinstance(value, datetime):
                data[key] = value.isoformat()
        return data
