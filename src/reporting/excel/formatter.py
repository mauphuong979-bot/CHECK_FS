from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP

from openpyxl.cell.cell import MergedCell
from openpyxl.comments import Comment

from src.reporting.excel import (
    CHECK_ROUND_QUANTUM,
)
from src.reporting.excel_styles import (
    WARNING_FILL,
)


def _excel_check_formula(expression: str) -> str:
    return f"=ROUND({expression},2)"


def _rounded_check_value(value: Decimal | float | int) -> Decimal:
    return Decimal(str(value)).quantize(CHECK_ROUND_QUANTUM, rounding=ROUND_HALF_UP)


def _has_check_difference(value: Decimal | float | int) -> bool:
    return _rounded_check_value(value) != Decimal(0)


def _format_vietnamese_number(value: Decimal | float | int) -> str:
    rounded = _rounded_check_value(abs(value))
    formatted = f"{rounded:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return formatted[:-3] if formatted.endswith(",00") else formatted


def _arithmetic_difference_message(value: Decimal | float | int) -> str:
    return f"Kiểm tra số học: chênh lệch {_format_vietnamese_number(value)}."


def _mark_arithmetic_difference(cell, value: Decimal | float | int) -> None:
    cell = _merged_anchor_cell(cell)
    cell.fill = WARNING_FILL
    cell.comment = Comment(_arithmetic_difference_message(value), "CHECK_FS_RULE")


def _merged_anchor_cell(cell):
    """Trả về ô neo có thể ghi của một vùng merge, nếu ``cell`` là ô phụ."""
    if not isinstance(cell, MergedCell):
        return cell
    for merged_range in cell.parent.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return cell.parent.cell(
                row=merged_range.min_row,
                column=merged_range.min_col,
            )
    return cell
