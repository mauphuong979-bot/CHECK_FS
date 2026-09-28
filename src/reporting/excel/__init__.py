from __future__ import annotations

from decimal import Decimal

from src.extraction.table_values import DEFAULT_EXCEL_DATA_START_ROW
from src.reporting.excel.formatter import (
    _excel_check_formula,
    _rounded_check_value,
    _has_check_difference,
    _format_vietnamese_number,
    _arithmetic_difference_message,
    _mark_arithmetic_difference,
    _merged_anchor_cell,
)
from src.reporting.excel.formulas import _build_sum_formula
from src.reporting.excel.reconciler import _ordered_note_matches

NOTE_STATUS_LABELS = {
    "Matched": "Khớp",
    "Difference": "Có chênh lệch",
    "Note value not found": "Không tìm thấy số liệu TM",
    "Statement not found": "Không tìm thấy chỉ tiêu BCTC",
}
DASH_ZERO_MARKERS = {"-", "–", "—"}
TITLE_ROW = 1
INFO_ROW = 2
HEADER_ROW = 3
DATA_START_ROW = DEFAULT_EXCEL_DATA_START_ROW
SUMMARY_SORT_TABLE_INDEX = "table_index_asc"
SUMMARY_SORT_ISSUES_DESC = "total_issues_desc"
SUMMARY_SORT_STATUS = "status_asc"
NOTE_SORT_TABLE_INDEX = "note_table_index_asc"
NOTE_SORT_STATUS = "status_asc"
NOTE_SORT_SCORE = "match_score_asc"
VALID_SUMMARY_SORTS = {
    SUMMARY_SORT_TABLE_INDEX,
    SUMMARY_SORT_ISSUES_DESC,
    SUMMARY_SORT_STATUS,
}
CHECK_ROUND_DIGITS = 2
CHECK_ROUND_QUANTUM = Decimal("0.01")
ACCOUNTING_NUMBER_FORMAT = '#,##0;(#,##0);-'
CHECK_DIFFERENCE_NUMBER_FORMAT = ACCOUNTING_NUMBER_FORMAT
