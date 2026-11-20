from __future__ import annotations

from decimal import Decimal


def to_decimal(value: Decimal | float | int) -> Decimal:
    if isinstance(value, Decimal):
        return value
    return Decimal(str(value))


def format_vietnamese_number(value: Decimal | float | int) -> str:
    """Định dạng số tuyệt đối kiểu Việt Nam, giữ đủ phần thập phân khác 0 (không làm tròn)."""
    number = abs(to_decimal(value))
    integer_part, _, fraction = f"{number:f}".partition(".")
    grouped = f"{int(integer_part):,}".replace(",", ".")
    fraction = fraction.rstrip("0")
    return f"{grouped},{fraction}" if fraction else grouped


def format_amount(value: Decimal | float | int) -> str:
    """Định dạng số có dấu theo kiểu kế toán: số âm đặt trong ngoặc."""
    number = to_decimal(value)
    formatted = format_vietnamese_number(number)
    return f"({formatted})" if number < 0 else formatted
