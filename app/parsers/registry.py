"""Minimal parser registry.

The knowledge-base ingestion pipeline in the host system asks the registry
for a parser by MIME type or file extension. Registering a parser means:
  registry.register(DocxParser())
after which registry.for_mime("application/vnd...wordprocessingml.document")
returns that parser instance.
"""

from __future__ import annotations

from typing import Callable, Dict, Iterable, List, Optional


class ParserRegistry:
    def __init__(self) -> None:
        self._by_mime: Dict[str, object] = {}
        self._by_ext: Dict[str, object] = {}
        self._order: List[object] = []

    def register(self, parser: object) -> object:
        for mime in getattr(parser, "mime_types", ()):
            self._by_mime[mime.lower()] = parser
        for ext in getattr(parser, "extensions", ()):
            self._by_ext[ext.lower().lstrip(".")] = parser
        if parser not in self._order:
            self._order.append(parser)
        return parser

    def for_mime(self, mime: str) -> Optional[object]:
        return self._by_mime.get((mime or "").lower())

    def for_extension(self, ext: str) -> Optional[object]:
        return self._by_ext.get((ext or "").lower().lstrip("."))

    def for_path(self, path: str) -> Optional[object]:
        name = str(path).rsplit("/", 1)[-1].rsplit("\\", 1)[-1]
        if "." not in name:
            return None
        return self.for_extension(name.rsplit(".", 1)[-1])

    def parsers(self) -> List[object]:
        return list(self._order)

    def mime_types(self) -> List[str]:
        return sorted(self._by_mime)


def build_default_registry(*, register: Iterable[Callable[[], object]] = ()) -> ParserRegistry:
    """Build a registry with the DOCX parser registered out of the box."""
    from app.parsers.docx_parser import DocxParser

    registry = ParserRegistry()
    registry.register(DocxParser())
    for factory in register:
        registry.register(factory())
    return registry
