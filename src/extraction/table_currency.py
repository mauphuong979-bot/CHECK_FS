from __future__ import annotations

import re

from src.domain.models import ExtractedTable, StatementType
from src.normalization.text_cleaner import normalized_key


CURRENCY_HEADER_ROWS = 4
# Thứ tự có ý nghĩa: đơn vị bội số (triệu/nghìn đồng) phải được nhận diện trước "đồng".
_CURRENCY_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("Triệu VND", re.compile(r"\btrieu (?:dong|vnd)\b")),
    ("Nghìn VND", re.compile(r"\b(?:nghin|ngan) (?:dong|vnd)\b")),
    ("USD", re.compile(r"\busd\b")),
    ("EUR", re.compile(r"\beur\b")),
    ("JPY", re.compile(r"\bjpy\b")),
    ("VND", re.compile(r"\bvnd\b|\bdon vi tinh: dong\b")),
)


def detect_table_currency(table: ExtractedTable) -> str | None:
    """Nhận diện đơn vị tiền tệ trình bày của bảng từ header và dòng "Đơn vị tính".

    Chỉ trả về kết quả khi có đúng một đơn vị được nhận diện; nhiều đơn vị
    hoặc không có căn cứ đều trả ``None`` để caller không suy diễn.
    """
    header_text = " ".join(
        value for row in table.rows[:CURRENCY_HEADER_ROWS] for value in row if value
    )
    found = _currencies_in(normalized_key(header_text))
    if not found:
        found = _currencies_in(normalized_key(table.title_hint))
    return next(iter(found)) if len(found) == 1 else None


def resolve_table_currencies(tables: list[ExtractedTable]) -> dict[int, str | None]:
    """Đơn vị tiền tệ theo bảng; bảng Thuyết minh không ghi đơn vị kế thừa bảng liền trước.

    Đơn vị tính thường chỉ ghi một lần cho cả phần Thuyết minh, nên bảng không có
    dòng đơn vị lấy theo bảng gần nhất phía trước đã nhận diện được.
    """
    currencies: dict[int, str | None] = {}
    last_known: str | None = None
    for table in tables:
        currency = detect_table_currency(table)
        if currency is None and table.statement_type == StatementType.NOTE:
            currency = last_known
        if currency is not None:
            last_known = currency
        currencies[table.index] = currency
    return currencies


def currencies_are_compatible(first: str | None, second: str | None) -> bool:
    return first is None or second is None or first == second


def _currencies_in(text: str) -> set[str]:
    found: set[str] = set()
    remaining = text
    for currency, pattern in _CURRENCY_PATTERNS:
        if pattern.search(remaining):
            found.add(currency)
            remaining = pattern.sub(" ", remaining)
    return found
