from __future__ import annotations

from decimal import Decimal, InvalidOperation
import re

from src.normalization.text_cleaner import clean_text


EMPTY_MARKERS = {"", "-", "–", "—", "n/a", "N/A"}


def parse_accounting_number(value: object) -> Decimal | None:
    text = clean_text(value)
    if text in EMPTY_MARKERS:
        return None

    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1].strip()

    text = text.replace("VND", "").replace("vnđ", "").replace("VNĐ", "")
    text = text.replace(" ", "")

    if text.startswith("-"):
        negative = True
        text = text[1:]

    text = _normalize_separators(text)
    if text is None or not re.fullmatch(r"\d+(\.\d+)?", text):
        return None

    try:
        number = Decimal(text)
    except InvalidOperation:
        return None

    return -number if negative else number


def looks_like_number(value: object) -> bool:
    return parse_accounting_number(value) is not None


def _normalize_separators(text: str) -> str | None:
    if "." in text and "," in text:
        if text.rfind(",") > text.rfind("."):
            return text.replace(".", "").replace(",", ".")
        return text.replace(",", "")

    if "." in text:
        parts = text.split(".")
        if len(parts) > 2 and all(len(part) == 3 for part in parts[1:]):
            return "".join(parts)
        if len(parts) == 2 and len(parts[1]) == 3 and len(parts[0]) <= 3:
            return "".join(parts)
        return text

    if "," in text:
        parts = text.split(",")
        if len(parts) > 2 and all(len(part) == 3 for part in parts[1:]):
            return "".join(parts)
        if len(parts) == 2 and len(parts[1]) == 3 and len(parts[0]) <= 3:
            return "".join(parts)
        return text.replace(",", ".")

    return text
