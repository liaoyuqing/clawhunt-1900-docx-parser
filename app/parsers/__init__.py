from app.parsers.docx_parser import DocxParser, parse_docx
from app.parsers.registry import ParserRegistry, build_default_registry

__all__ = ["DocxParser", "parse_docx", "ParserRegistry", "build_default_registry"]
