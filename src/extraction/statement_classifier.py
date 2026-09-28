from __future__ import annotations

import re

from src.domain.models import StatementType
from src.normalization.text_cleaner import clean_text, normalized_key


def _joined(rows: list[list[str]], max_rows: int = 8) -> str:
    return " ".join(
        normalized_key(cell)
        for row in rows[:max_rows]
        for cell in row
        if cell
    )


def get_title_hint(rows: list[list[str]]) -> str:
    for row in rows[:8]:
        values = [cell.strip() for cell in row if cell and cell.strip()]
        if values:
            return " | ".join(values)[:120]
    return ""


def _looks_like_informational_note(text: str) -> bool:
    if "so nam khau hao" in text or "thoi gian khau hao" in text:
        return True
    has_current_period = "so cuoi nam" in text or "nam nay" in text
    has_prior_period = "so dau nam" in text or "nam truoc" in text
    return has_current_period and has_prior_period


def _looks_like_depreciation_life_range_table(
    rows: list[list[str]],
    title_hint: str,
    context_hints: tuple[str, ...],
) -> bool:
    has_year_header = any(
        normalized_key(value) == "nam"
        for row in rows[:3]
        for value in row
    )
    range_count = sum(
        1
        for row in rows
        for value in row[1:]
        if re.fullmatch(r"\d+\s*[-–—]\s*\d+", clean_text(value))
    )
    asset_class_count = sum(
        1
        for row in rows
        if row
        and any(
            marker in normalized_key(row[0])
            for marker in (
                "nha cua",
                "may moc",
                "phuong tien van tai",
                "thiet bi van phong",
            )
        )
    )
    context = normalized_key(" ".join((title_hint, *context_hints)))
    has_depreciation_context = "khau hao" in context
    return (
        has_year_header
        and asset_class_count >= 2
        and (range_count >= 2 or has_depreciation_context)
    )


def _looks_like_related_party_relationship_table(
    rows: list[list[str]],
    title_hint: str,
    context_hints: tuple[str, ...],
) -> bool:
    context = normalized_key(" ".join((title_hint, *context_hints)))
    header = _joined(rows, max_rows=3)
    return (
        "giao dich voi cac ben lien quan" in context
        and "ben lien quan" in header
        and "moi quan he" in header
    )


def _looks_like_table_of_contents(
    rows: list[list[str]],
    title_hint: str,
    context_hints: tuple[str, ...],
) -> bool:
    context = normalized_key(" ".join((title_hint, *context_hints)))
    header = _joined(rows, max_rows=2)
    return "muc luc" in context or ("noi dung" in header and "trang" in header)


def _looks_like_management_information(
    rows: list[list[str]],
    title_hint: str,
    context_hints: tuple[str, ...],
) -> bool:
    context = normalized_key(" ".join((title_hint, *context_hints)))
    header = _joined(rows, max_rows=2)
    return (
        "cac thanh vien ban giam doc" in context
        or "danh sach ban giam doc" in context
        or "ban giam doc" in header
    )


def _looks_like_business_location_information(
    title_hint: str,
    context_hints: tuple[str, ...],
) -> bool:
    context = normalized_key(" ".join((title_hint, *context_hints)))
    return "dia diem kinh doanh" in context


def _looks_like_signature_block(
    rows: list[list[str]],
    title_hint: str,
    context_hints: tuple[str, ...],
) -> bool:
    context = normalized_key(" ".join((title_hint, *context_hints)))
    text = _joined(rows)
    return (
        "thay mat va dai dien" in context
        or "tuq giam doc" in text
        or "tuq. giam doc" in text
    )


def classify_table(
    rows: list[list[str]],
    *,
    title_hint: str = "",
    context_hints: tuple[str, ...] = (),
) -> StatementType:
    text = _joined(rows)

    if _looks_like_related_party_relationship_table(rows, title_hint, context_hints):
        return StatementType.NOTE
    if _looks_like_table_of_contents(rows, title_hint, context_hints):
        return StatementType.GENERAL_INFO
    if _looks_like_management_information(rows, title_hint, context_hints):
        return StatementType.GENERAL_INFO
    if _looks_like_business_location_information(title_hint, context_hints):
        return StatementType.GENERAL_INFO
    if _looks_like_signature_block(rows, title_hint, context_hints):
        return StatementType.SIGNATURE
    if _looks_like_depreciation_life_range_table(rows, title_hint, context_hints):
        return StatementType.NOTE
    if "cong ty" in text and ("ma so thue" in text or "bao cao mau" in text):
        return StatementType.GENERAL_INFO
    if "luu chuyen tien" in text:
        return StatementType.CASH_FLOW
    if "tai san" in text and "ma so" in text:
        return StatementType.BALANCE_SHEET_ASSETS
    if "nguon von" in text and "ma so" in text:
        return StatementType.BALANCE_SHEET_EQUITY
    if "chi tieu" in text and "doanh thu" in text and "ma so" in text:
        return StatementType.INCOME_STATEMENT
    if "tong giam doc" in text or "nguoi lap bieu" in text or "ke toan truong" in text:
        return StatementType.SIGNATURE
    if (
        "vnd" in text
        or "31/12" in text
        or "01/01" in text
        or _looks_like_informational_note(text)
    ):
        return StatementType.NOTE
    labels = {
        normalized_key(row[0])
        for row in rows
        if row and row[0]
    }
    context = normalized_key(" ".join((title_hint, *context_hints)))
    if {"ngan han", "dai han"}.issubset(labels) and context:
        return StatementType.NOTE
    if (
        any(label.startswith("du phong phai thu") for label in labels)
        and any(label.startswith("gia tri thuan") for label in labels)
        and context
    ):
        return StatementType.NOTE
    return StatementType.UNKNOWN
