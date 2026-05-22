from __future__ import annotations

from html import unescape
from html.parser import HTMLParser


class _TextParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        self.parts.append(data)


class BeautifulSoup:
    def __init__(self, markup: str | None, parser: str | None = None) -> None:
        self._parser = _TextParser()
        self._parser.feed(markup or "")

    def get_text(self, separator: str = "", strip: bool = False) -> str:
        parts = [part.strip() if strip else part for part in self._parser.parts]
        parts = [part for part in parts if part] if strip else parts
        return unescape(separator.join(parts))
