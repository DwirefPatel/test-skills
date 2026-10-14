"""Parser interface, broker registry and shared text helpers."""

from __future__ import annotations

import re
import shutil
import subprocess
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Optional

from ..models import Package


class ParseError(Exception):
    """Raised when the text cannot be mapped reliably. Never guess or plug."""


def extract_pages(pdf_path) -> list[str]:
    """Run ``pdftotext -layout`` and return one string per page."""
    if not shutil.which("pdftotext"):
        raise ParseError("pdftotext (poppler-utils) is required to read PDFs")
    proc = subprocess.run(["pdftotext", "-layout", str(pdf_path), "-"],
                          capture_output=True, text=True, check=False)
    if proc.returncode != 0:
        raise ParseError(f"pdftotext failed: {proc.stderr.strip()}")
    pages = proc.stdout.split("\f")
    if pages and not pages[-1].strip():
        pages = pages[:-1]
    return pages


# CUSIP: 9 chars, last is a check digit. Validating it keeps random tokens out.
CUSIP_RE = re.compile(r"\b([0-9]{3}[0-9A-Z]{5}[0-9])\b")


def cusip_valid(cusip: str) -> bool:
    if len(cusip) != 9:
        return False
    total = 0
    for i, ch in enumerate(cusip[:8]):
        if ch.isdigit():
            v = int(ch)
        elif ch.isalpha():
            v = ord(ch.upper()) - 55
        elif ch == "*":
            v = 36
        elif ch == "@":
            v = 37
        elif ch == "#":
            v = 38
        else:
            return False
        if i % 2 == 1:
            v *= 2
        total += v // 10 + v % 10
    return (10 - total % 10) % 10 == int(cusip[8]) if cusip[8].isdigit() else False


def find_cusip(line: str) -> Optional[str]:
    for m in CUSIP_RE.finditer(line):
        if cusip_valid(m.group(1)):
            return m.group(1)
    return None


EIN_RE = re.compile(r"\b(\d{2}-\d{7})\b")
AMOUNT = r"\(?-?\$?[\d,]*\d\.\d{2}\)?-?"
AMOUNT_RE = re.compile(rf"^{AMOUNT}$")


class BrokerParser(ABC):
    """Subclass per broker layout and register with :func:`register`."""

    name: str = "base"
    #: lowercase strings; any one found in the first pages selects this parser
    detect_keywords: tuple[str, ...] = ()

    @classmethod
    def detect(cls, pages: list[str]) -> bool:
        head = "\n".join(pages[:3]).lower()
        return any(k in head for k in cls.detect_keywords)

    @abstractmethod
    def parse(self, pages: list[str], source_file: str = "") -> Package: ...


_REGISTRY: list[type[BrokerParser]] = []


def register(cls: type[BrokerParser]) -> type[BrokerParser]:
    _REGISTRY.insert(0, cls)  # later registrations take priority over the generic parser
    return cls


def registered() -> list[type[BrokerParser]]:
    return list(_REGISTRY)


def get_parser(pages: list[str], broker: Optional[str] = None) -> BrokerParser:
    if broker:
        for cls in _REGISTRY:
            if cls.name == broker.lower():
                return cls()
        raise ParseError(f"No parser registered for broker {broker!r}; available: "
                         f"{[c.name for c in _REGISTRY]}")
    for cls in _REGISTRY:
        if cls.detect(pages):
            return cls()
    raise ParseError("Could not detect the broker layout; pass broker=...")


def parse_pdf(pdf_path, broker: Optional[str] = None) -> Package:
    pages = extract_pages(pdf_path)
    return get_parser(pages, broker).parse(pages, source_file=Path(pdf_path).name)


def parse_text(pages: list[str], broker: Optional[str] = None, source_file: str = "") -> Package:
    return get_parser(pages, broker).parse(pages, source_file=source_file)
