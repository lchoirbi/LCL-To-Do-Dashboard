from __future__ import annotations

import calendar
import re
from datetime import date, datetime, timedelta


MONTHS = {name.lower(): num for num, name in enumerate(calendar.month_name) if name}
MONTHS.update({name.lower(): num for num, name in enumerate(calendar.month_abbr) if name})
WEEKDAYS = {name.lower(): num for num, name in enumerate(calendar.day_name)}
WEEKDAYS.update({name.lower(): num for num, name in enumerate(calendar.day_abbr)})


def _base(settings: dict | None = None) -> datetime:
    value = (settings or {}).get("RELATIVE_BASE")
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    return datetime.now()


def _safe_date(year: int, month: int, day: int) -> datetime | None:
    try:
        return datetime(year, month, day)
    except ValueError:
        return None


def parse(value: str | None, settings: dict | None = None) -> datetime | None:
    if not value:
        return None
    text = str(value).strip()
    normalized = re.sub(r"(\.\d{6})\d+", r"\1", text.replace("Z", "+00:00"))
    try:
        return datetime.fromisoformat(normalized)
    except ValueError:
        pass

    base = _base(settings)
    lower = text.lower()
    if re.fullmatch(r"today|eod|cob", lower):
        return base
    if lower == "tomorrow":
        return base + timedelta(days=1)

    match = re.search(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b", lower)
    if match:
        return _safe_date(*(int(part) for part in match.groups()))

    match = re.search(r"\b(\d{1,2})/(\d{1,2})/(\d{2,4})\b", lower)
    if match:
        month, day, year = (int(part) for part in match.groups())
        return _safe_date(2000 + year if year < 100 else year, month, day)

    month_names = "|".join(sorted(MONTHS, key=len, reverse=True))
    match = re.search(rf"\b({month_names})\.?\s+(\d{{1,2}})(?:st|nd|rd|th)?(?:,\s*(\d{{4}}))?\b", lower)
    if match:
        month = MONTHS[match.group(1).rstrip(".")]
        day = int(match.group(2))
        year = int(match.group(3) or base.year)
        candidate = _safe_date(year, month, day)
        if candidate and candidate.date() < base.date() - timedelta(days=7) and not match.group(3):
            candidate = _safe_date(year + 1, month, day)
        return candidate

    for word, weekday in WEEKDAYS.items():
        if re.fullmatch(rf"(next\s+)?{word}", lower):
            days = (weekday - base.weekday()) % 7
            if days == 0 or lower.startswith("next "):
                days += 7
            return base + timedelta(days=days)
    return None
