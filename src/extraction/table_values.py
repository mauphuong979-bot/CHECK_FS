from __future__ import annotations

from decimal import Decimal

from src.domain.models import (
    ExtractedTable,
    TableValue,
    TableValueStatus,
)
from src.extraction.table_layout import (
    TableLayout,
    detect_table_layout,
)
from src.normalization.number_parser import parse_accounting_number
from src.normalization.text_cleaner import clean_text


DEFAULT_EXCEL_DATA_START_ROW = 4
DASH_ZERO_MARKERS = {"-", "–", "—"}
MISSING_VALUE_MARKERS = {"", "n/a"}


def normalize_statement_code(value: object) -> str:
    text = clean_text(value)
    if not text:
        return ""
    parsed_number = parse_accounting_number(text)
    if parsed_number is not None and parsed_number == parsed_number.to_integral_value():
        return str(int(parsed_number))
    return text.lower()


def code_excel_row_lookup(
    table: ExtractedTable,
    layout: TableLayout | None = None,
    *,
    data_start_row: int = DEFAULT_EXCEL_DATA_START_ROW,
) -> dict[str, int]:
    resolved_layout = layout or detect_table_layout(table)
    code_rows: dict[str, int] = {}
    for excel_row, row in enumerate(table.rows, start=data_start_row):
        if len(row) < resolved_layout.code_col:
            continue
        code = normalize_statement_code(row[resolved_layout.code_col - 1])
        if code:
            code_rows[code] = excel_row
    return code_rows


def table_value_lookup(
    table: ExtractedTable,
    layout: TableLayout,
) -> dict[str, dict[int, float]]:
    typed_values = table_value_state_lookup(table, layout)
    return {
        code: {
            column: float(cell_value.value) if cell_value.value is not None else 0.0
            for column, cell_value in values_by_column.items()
        }
        for code, values_by_column in typed_values.items()
    }


def table_value_state_lookup(
    table: ExtractedTable,
    layout: TableLayout,
) -> dict[str, dict[int, TableValue]]:
    table_values: dict[str, dict[int, TableValue]] = {}
    for row in table.rows:
        if len(row) < layout.code_col:
            continue
        code = normalize_statement_code(row[layout.code_col - 1])
        if not code:
            continue
        values_by_column: dict[int, TableValue] = {}
        for value_column in (layout.current_col, layout.prior_col):
            is_present = len(row) >= value_column
            raw_value = row[value_column - 1] if is_present else ""
            values_by_column[value_column] = classify_table_value(
                raw_value,
                is_present=is_present,
            )
        table_values[code] = values_by_column
    return table_values


def classify_table_value(value: object, *, is_present: bool = True) -> TableValue:
    if not is_present:
        return TableValue(TableValueStatus.MISSING)

    text = clean_text(value)
    if text in DASH_ZERO_MARKERS:
        return TableValue(TableValueStatus.DASH_ZERO, Decimal(0))
    if text.casefold() in MISSING_VALUE_MARKERS:
        return TableValue(TableValueStatus.MISSING)

    parsed_number = parse_accounting_number(text)
    if parsed_number is None:
        return TableValue(TableValueStatus.INVALID)
    if parsed_number == 0:
        return TableValue(TableValueStatus.ZERO, parsed_number)
    return TableValue(TableValueStatus.NUMBER, parsed_number)
