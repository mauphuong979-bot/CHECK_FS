from __future__ import annotations

from decimal import Decimal
from src.reporting.excel.formatter import _excel_check_formula, _has_check_difference, _mark_arithmetic_difference, _merged_anchor_cell


def _build_sum_formula(cell_references: list[str]) -> str:
    if not cell_references:
        return ""
    return f"=SUM({','.join(cell_references)})"
