from decimal import Decimal
import unittest

from src.domain.models import (
    ExtractedTable,
    StatementType,
    TableValue,
    TableValueStatus,
)
from src.extraction.table_layout import TableLayout
from src.extraction.table_values import (
    code_excel_row_lookup,
    classify_table_value,
    normalize_statement_code,
    table_value_lookup,
    table_value_state_lookup,
)


class TableValueLookupTest(unittest.TestCase):
    def test_normalizes_numeric_and_text_statement_codes(self):
        self.assertEqual(normalize_statement_code("01"), "1")
        self.assertEqual(normalize_statement_code(" 421A "), "421a")
        self.assertEqual(normalize_statement_code(""), "")

    def test_maps_codes_to_excel_rows_using_current_data_offset(self):
        table = ExtractedTable(
            1,
            [
                ["Chỉ tiêu", "Mã số", "Năm nay", "Năm trước"],
                ["Tiền", "110", "100", "90"],
                ["Tổng cộng", "110", "200", "180"],
            ],
            StatementType.BALANCE_SHEET_ASSETS,
        )

        self.assertEqual(
            code_excel_row_lookup(table, TableLayout(2, 3, 4)),
            {"mã số": 4, "110": 6},
        )

    def test_allows_explicit_excel_data_start_row(self):
        table = ExtractedTable(
            1,
            [["Chỉ tiêu", "110", "100", "90"]],
            StatementType.BALANCE_SHEET_ASSETS,
        )

        self.assertEqual(
            code_excel_row_lookup(
                table,
                TableLayout(2, 3, 4),
                data_start_row=10,
            ),
            {"110": 10},
        )

    def test_preserves_numeric_values_as_float_for_reporter_compatibility(self):
        table = ExtractedTable(
            1,
            [["Tiền", "110", "1.000", "(900)"]],
            StatementType.BALANCE_SHEET_ASSETS,
        )

        self.assertEqual(
            table_value_lookup(table, TableLayout(2, 3, 4)),
            {"110": {3: 1000.0, 4: -900.0}},
        )

    def test_preserves_zero_fallback_for_dash_blank_invalid_and_missing_values(self):
        table = ExtractedTable(
            1,
            [
                ["Dấu gạch", "110", "-", "—"],
                ["Ô trống", "120", "", ""],
                ["Không hợp lệ", "130", "abc", "n/a"],
                ["Thiếu cột", "140"],
            ],
            StatementType.BALANCE_SHEET_ASSETS,
        )

        self.assertEqual(
            table_value_lookup(table, TableLayout(2, 3, 4)),
            {
                "110": {3: 0.0, 4: 0.0},
                "120": {3: 0.0, 4: 0.0},
                "130": {3: 0.0, 4: 0.0},
                "140": {3: 0.0, 4: 0.0},
            },
        )

    def test_classifies_number_and_explicit_zero_without_losing_decimal(self):
        self.assertEqual(
            classify_table_value("1.234,50"),
            TableValue(TableValueStatus.NUMBER, Decimal("1234.50")),
        )
        self.assertEqual(
            classify_table_value("0.00"),
            TableValue(TableValueStatus.ZERO, Decimal("0.00")),
        )

    def test_classifies_dash_missing_and_invalid_as_distinct_states(self):
        self.assertEqual(
            classify_table_value("—"),
            TableValue(TableValueStatus.DASH_ZERO, Decimal(0)),
        )
        self.assertEqual(
            classify_table_value(""),
            TableValue(TableValueStatus.MISSING),
        )
        self.assertEqual(
            classify_table_value("N/A"),
            TableValue(TableValueStatus.MISSING),
        )
        self.assertEqual(
            classify_table_value("abc"),
            TableValue(TableValueStatus.INVALID),
        )
        self.assertEqual(
            classify_table_value("", is_present=False),
            TableValue(TableValueStatus.MISSING),
        )

    def test_typed_lookup_retains_cell_state_by_period_column(self):
        table = ExtractedTable(
            1,
            [
                ["Số hợp lệ", "110", "100", "0"],
                ["Dấu gạch", "120", "-", "abc"],
                ["Thiếu cột", "130"],
            ],
            StatementType.BALANCE_SHEET_ASSETS,
        )

        self.assertEqual(
            table_value_state_lookup(table, TableLayout(2, 3, 4)),
            {
                "110": {
                    3: TableValue(TableValueStatus.NUMBER, Decimal("100")),
                    4: TableValue(TableValueStatus.ZERO, Decimal("0")),
                },
                "120": {
                    3: TableValue(TableValueStatus.DASH_ZERO, Decimal("0")),
                    4: TableValue(TableValueStatus.INVALID),
                },
                "130": {
                    3: TableValue(TableValueStatus.MISSING),
                    4: TableValue(TableValueStatus.MISSING),
                },
            },
        )


if __name__ == "__main__":
    unittest.main()
