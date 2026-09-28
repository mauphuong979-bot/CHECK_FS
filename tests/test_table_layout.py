import unittest

from src.domain.models import ExtractedTable, StatementType
from src.extraction.table_layout import (
    TableLayout,
    detect_code_column,
    detect_period_columns_from_headers,
    detect_table_layout,
    is_note_column,
    numeric_columns_for,
    select_period_column,
)


class TableLayoutTest(unittest.TestCase):
    def test_detects_standard_statement_columns(self):
        table = ExtractedTable(
            1,
            [
                ["Chỉ tiêu", "Mã số", "Thuyết minh", "Năm nay", "Năm trước"],
                ["Tiền", "110", "V.01", "1.000", "900"],
            ],
            StatementType.BALANCE_SHEET_ASSETS,
        )

        self.assertEqual(detect_table_layout(table), TableLayout(2, 4, 5))
        self.assertEqual(numeric_columns_for(table), {2, 4, 5})
        self.assertTrue(is_note_column(table, 3))

    def test_uses_non_note_candidates_when_data_has_no_numbers(self):
        table = ExtractedTable(
            1,
            [["Chỉ tiêu", "Mã / số", "Thuyết minh", "Kỳ này", "Kỳ trước"]],
            StatementType.BALANCE_SHEET_ASSETS,
        )

        self.assertEqual(detect_table_layout(table), TableLayout(2, 4, 5))

    def test_uses_same_period_column_when_only_one_candidate_exists(self):
        table = ExtractedTable(
            1,
            [["Chỉ tiêu", "Mã số", "Giá trị"]],
            StatementType.INCOME_STATEMENT,
        )

        self.assertEqual(detect_table_layout(table), TableLayout(2, 3, 3))

    def test_falls_back_safely_for_single_column_table(self):
        table = ExtractedTable(
            1,
            [["Nội dung"]],
            StatementType.UNKNOWN,
        )

        self.assertEqual(detect_code_column(table), 1)
        self.assertEqual(detect_table_layout(table), TableLayout(1, 1, 1))

    def test_code_header_search_is_limited_to_first_eight_rows(self):
        table = ExtractedTable(
            1,
            [["", ""] for _ in range(8)] + [["Mã số", "Giá trị"]],
            StatementType.UNKNOWN,
        )

        self.assertEqual(detect_code_column(table), 2)

    def test_selects_period_from_source_statement_layout(self):
        current_layout = TableLayout(2, 4, 5)
        source_layout = TableLayout(2, 6, 7)
        layouts = {StatementType.INCOME_STATEMENT: source_layout}

        self.assertEqual(
            select_period_column(
                StatementType.CASH_FLOW,
                StatementType.CASH_FLOW,
                "current",
                current_layout,
                layouts,
            ),
            4,
        )
        self.assertEqual(
            select_period_column(
                StatementType.CASH_FLOW,
                StatementType.INCOME_STATEMENT,
                "prior",
                current_layout,
                layouts,
            ),
            7,
        )

    def test_detects_period_columns_from_date_range_years(self):
        table = ExtractedTable(
            1,
            [
                ["", "Từ 01/01/2025 đến 31/12/2025", "", "Từ 14/05/2024 đến 31/12/2024"],
                ["", "VND", "", "VND"],
                ["Lợi nhuận trước thuế", "100", "", "90"],
            ],
            StatementType.NOTE,
        )

        self.assertEqual(detect_period_columns_from_headers(table), ([1], [3]))


if __name__ == "__main__":
    unittest.main()
