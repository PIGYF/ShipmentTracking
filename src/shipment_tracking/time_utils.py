from __future__ import annotations

from datetime import datetime, timezone, timedelta


CHINA_TZ = timezone(timedelta(hours=8))


def to_china_naive(value: datetime | None) -> datetime | None:
    if value is None:
        return None
    if value.tzinfo is None:
        return value.replace(tzinfo=None)
    return value.astimezone(CHINA_TZ).replace(tzinfo=None)


def date_only(value: datetime | None) -> datetime | None:
    value = to_china_naive(value)
    if value is None:
        return None
    return value.replace(hour=0, minute=0, second=0, microsecond=0)
