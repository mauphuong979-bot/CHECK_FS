from __future__ import annotations

from dataclasses import dataclass
import re

from src.domain.models import ExtractedTable, StatementType
from src.normalization.number_parser import parse_accounting_number
from src.normalization.text_cleaner import normalized_key


@dataclass(frozen=True)
class TableLayout:
    code_col: int
    current_col: int
    prior_col: int


def detect_period_columns_from_headers(
    table: ExtractedTable,
) -> tuple[list[int], list[int]]:
    """Nhận diện cột kỳ hiện tại/kỳ trước từ nhãn hoặc năm trong header (0-based)."""
    current_cols: list[int] = []
    prior_cols: list[int] = []
    years_by_col: dict[int, set[int]] = {}
    for col in range(table.column_count):
        header = normalized_key(
            " ".join(row[col] for row in table.rows[:3] if len(row) > col)
        )
        if "so cuoi nam" in header or "nam nay" in header:
            current_cols.append(col)
        if "so dau nam" in header or "nam truoc" in header:
            prior_cols.append(col)
        years = {int(value) for value in re.findall(r"\b(?:19|20)\d{2}\b", header)}
        if years:
            years_by_col[col] = years

    if current_cols or prior_cols:
        return current_cols, prior_cols

    distinct_years = {year for years in years_by_col.values() for year in years}
    if len(distinct_years) < 2:
        return [], []
    latest_year = max(distinct_years)
    earliest_year = min(distinct_years)
    return (
        [col for col, years in years_by_col.items() if latest_year in years],
        [col for col, years in years_by_col.items() if earliest_year in years],
    )


def detect_table_layout(table: ExtractedTable) -> TableLayout:
    code_col = detect_code_column(table)
    numeric_columns = sorted(
        column for column in numeric_columns_for(table) if column > code_col
    )
    if len(numeric_columns) >= 2:
        return TableLayout(
            code_col=code_col,
            current_col=numeric_columns[0],
            prior_col=numeric_columns[1],
        )

    candidate_columns = [
        column
        for column in range(code_col + 1, table.column_count + 1)
        if not is_note_column(table, column)
    ]
    if len(candidate_columns) >= 2:
        return TableLayout(
            code_col=code_col,
            current_col=candidate_columns[0],
            prior_col=candidate_columns[1],
        )
    if candidate_columns:
        return TableLayout(
            code_col=code_col,
            current_col=candidate_columns[0],
            prior_col=candidate_columns[0],
        )

    fallback_column = min(code_col + 1, max(table.column_count, code_col))
    return TableLayout(
        code_col=code_col,
        current_col=fallback_column,
        prior_col=fallback_column,
    )


def detect_code_column(table: ExtractedTable) -> int:
    for row in table.rows[:8]:
        for column, value in enumerate(row, start=1):
            key = normalized_key(value)
            if key in {"ma so", "ma/so", "ma / so"} or ("ma" in key and "so" in key):
                return column
    return 2 if table.column_count >= 2 else 1


def is_note_column(table: ExtractedTable, column: int) -> bool:
    for row in table.rows[:8]:
        if len(row) >= column and "thuyet minh" in normalized_key(row[column - 1]):
            return True
    return False


def numeric_columns_for(table: ExtractedTable) -> set[int]:
    numeric_columns: set[int] = set()
    for row in table.rows:
        for column, value in enumerate(row, start=1):
            if parse_accounting_number(value) is not None:
                numeric_columns.add(column)
    return numeric_columns


def select_period_column(
    current_type: StatementType,
    source_type: StatementType,
    period: str,
    current_layout: TableLayout,
    statement_layouts: dict[StatementType, TableLayout],
) -> int:
    layout = (
        current_layout
        if source_type == current_type
        else statement_layouts.get(source_type, current_layout)
    )
    return layout.current_col if period == "current" else layout.prior_col
