from __future__ import annotations

import re

from . import MONTHS, WEEKDAYS, parse


def search_dates(text: str, settings: dict | None = None):
    candidates: list[tuple[int, str]] = []
    month_names = "|".join(sorted(MONTHS, key=len, reverse=True))
    weekday_names = "|".join(sorted(WEEKDAYS, key=len, reverse=True))
    patterns = [
        r"\b\d{4}-\d{1,2}-\d{1,2}\b",
        r"\b\d{1,2}/\d{1,2}/\d{2,4}\b",
        rf"\b(?:{month_names})\.?\s+\d{{1,2}}(?:st|nd|rd|th)?(?:,\s*\d{{4}})?\b",
        rf"\b(?:next\s+)?(?:{weekday_names})\b",
        r"\b(?:today|tomorrow|eod|cob)\b",
    ]
    for pattern in patterns:
        for match in re.finditer(pattern, text or "", flags=re.IGNORECASE):
            candidates.append((match.start(), match.group(0)))

    found = []
    seen = set()
    for _, phrase in sorted(candidates):
        key = phrase.lower()
        if key in seen:
            continue
        seen.add(key)
        parsed = parse(phrase, settings=settings)
        if parsed:
            found.append((phrase, parsed))
    return found or None
