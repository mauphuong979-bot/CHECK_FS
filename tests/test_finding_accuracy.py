"""Hồi quy cho độ chính xác của phát hiện: không bỏ sót chênh lệch nhỏ, không tạo lỗi rác."""

from io import BytesIO
import unittest

from openpyxl import load_workbook

from src.domain.models import ExtractedTable, StatementType, WordCommentCategory
from src.extraction.table_currency import detect_table_currency, resolve_table_currencies
from src.reconciliation.note_matcher import reconcile_note_tables
from src.reporting.excel_report import (
    DECIMAL_ACCOUNTING_NUMBER_FORMAT,
    build_audit_workbook_with_comment_targets,
)


def _arithmetic_targets(targets):
    """Bảng TM đứng riêng không có BCTC để đối chiếu; chỉ xét phát hiện số học."""
    return [target for target in targets if not target.message.startswith("Đối chiếu")]


def _note(index: int, rows: list[list[str]], title: str = "Bảng thuyết minh") -> ExtractedTable:
    return ExtractedTable(index, rows, StatementType.NOTE, title)


class ExactArithmeticTest(unittest.TestCase):
    def test_one_cent_vertical_difference_is_reported_with_exact_amount(self):
        table = _note(
            1,
            [
                ["", "Số cuối năm"],
                ["", "USD"],
                ["Khoản A", "100,10"],
                ["Khoản B", "200,20"],
                ["Tổng cộng", "300,31"],
            ],
        )

        workbook_bytes, all_targets = build_audit_workbook_with_comment_targets([table])
        targets = _arithmetic_targets(all_targets)

        self.assertEqual(len(targets), 1)
        self.assertEqual(targets[0].category, WordCommentCategory.DIFFERENCE)
        self.assertEqual(
            targets[0].message,
            "Kiểm tra cộng dọc: số trình bày 300,31; số tính lại "
            "(cộng từ “Khoản A” đến “Khoản B”) 300,3; chênh lệch 0,01.",
        )
        worksheet = load_workbook(BytesIO(workbook_bytes))["T001_TM"]
        self.assertEqual(worksheet["B8"].number_format, DECIMAL_ACCOUNTING_NUMBER_FORMAT)
        check_formulas = [
            cell.value
            for row in worksheet.iter_rows()
            for cell in row
            if isinstance(cell.value, str) and cell.value.startswith("=ROUND(")
        ]
        self.assertIn("=ROUND(B8-SUM(B6:B7),2)", check_formulas)

    def test_matching_decimal_total_has_no_finding(self):
        table = _note(
            1,
            [["", "Số cuối năm"], ["Khoản A", "100,10"], ["Khoản B", "200,20"], ["Tổng cộng", "300,30"]],
        )

        _, targets = build_audit_workbook_with_comment_targets([table])

        self.assertEqual(_arithmetic_targets(targets), [])

    def test_excel_round_digits_follow_source_precision(self):
        table = _note(
            1,
            [["", "Số cuối năm"], ["Khoản A", "1,1255"], ["Khoản B", "2,2500"], ["Tổng cộng", "3,3756"]],
        )

        workbook_bytes, targets = build_audit_workbook_with_comment_targets([table])

        worksheet = load_workbook(BytesIO(workbook_bytes))["T001_TM"]
        formulas = [cell.value for row in worksheet.iter_rows() for cell in row if isinstance(cell.value, str)]
        self.assertIn("=ROUND(B7-SUM(B5:B6),4)", formulas)
        self.assertIn("chênh lệch 0,0001", _arithmetic_targets(targets)[0].message)


class DeductionSectionTest(unittest.TestCase):
    ROWS = [
        ["", "Số đầu năm", "Đã trả/Phân loại"],
        ["Vay dài hạn:", "", ""],
        ["Ngân hàng A", "330.000.000", "(120.000.000)"],
        ["Ngân hàng B", "551.000.000", "(114.000.000)"],
        ["Trừ:", "", ""],
        ["Nợ dài hạn đến hạn trả", "", ""],
        ["Ngân hàng A", "120.000.000", "(120.000.000)"],
        ["Ngân hàng B", "114.000.000", "(114.000.000)"],
    ]

    def test_total_after_deduction_header_is_base_minus_deduction(self):
        table = _note(1, [*self.ROWS, ["Tổng cộng", "647.000.000", "-"]], "Vay dài hạn")

        workbook_bytes, targets = build_audit_workbook_with_comment_targets([table])

        self.assertEqual(_arithmetic_targets(targets), [])
        worksheet = load_workbook(BytesIO(workbook_bytes))["T001_TM"]
        formulas = [cell.value for row in worksheet.iter_rows() for cell in row if isinstance(cell.value, str)]
        self.assertIn("=ROUND(B12-(SUM(B6:B7)-SUM(B9:B11)),2)", formulas)

    def test_real_error_in_deduction_layout_is_still_reported(self):
        table = _note(1, [*self.ROWS, ["Tổng cộng", "647.000.001", "-"]], "Vay dài hạn")

        _, all_targets = build_audit_workbook_with_comment_targets([table])
        targets = _arithmetic_targets(all_targets)

        self.assertEqual(len(targets), 1)
        self.assertIn("chênh lệch 1.", targets[0].message)

    def test_negative_presented_deduction_is_summed(self):
        table = _note(
            1,
            [
                ["", "Số cuối năm"],
                ["Phải thu khách hàng", "1.000"],
                ["Trừ:", ""],
                ["Dự phòng phải thu khó đòi", "(100)"],
                ["Tổng cộng", "900"],
            ],
        )

        _, targets = build_audit_workbook_with_comment_targets([table])

        self.assertEqual(_arithmetic_targets(targets), [])


class CurrencyMatchingTest(unittest.TestCase):
    @staticmethod
    def _balance_sheet(index: int, unit: str, total: str) -> ExtractedTable:
        return ExtractedTable(
            index,
            [
                ["TÀI SẢN", "Mã số", "Thuyết minh", "Số cuối năm", "Số đầu năm"],
                ["Hàng tồn kho", "140", "V.1", total, total],
            ],
            StatementType.BALANCE_SHEET_ASSETS,
            f"Đơn vị tính: {unit}",
        )

    def test_detects_currency_from_header_or_unit_line(self):
        self.assertEqual(detect_table_currency(self._balance_sheet(1, "USD", "1")), "USD")
        self.assertEqual(detect_table_currency(self._balance_sheet(1, "VND", "1")), "VND")
        self.assertEqual(
            detect_table_currency(_note(1, [["", "Số cuối năm"], ["", "Triệu đồng"]])),
            "Triệu VND",
        )
        self.assertIsNone(detect_table_currency(_note(1, [["Hoạt động kinh doanh", "1"]])))

    def test_note_without_unit_inherits_previous_table_currency(self):
        tables = [
            self._balance_sheet(1, "USD", "1"),
            _note(2, [["", "Số cuối năm"], ["Tổng cộng", "1"]]),
            self._balance_sheet(3, "VND", "1"),
        ]

        self.assertEqual(resolve_table_currencies(tables), {1: "USD", 2: "USD", 3: "VND"})

    def test_usd_note_is_reconciled_with_usd_statement_not_vnd_translation(self):
        tables = [
            self._balance_sheet(1, "USD", "1.000,50"),
            _note(
                2,
                [["", "Số cuối năm", "Số đầu năm"], ["", "USD", "USD"], ["Tổng cộng", "1.000,50", "1.000,50"]],
                "Hàng tồn kho",
            ),
            self._balance_sheet(3, "VND", "26.000.000.000"),
        ]

        matches = [match for match in reconcile_note_tables(tables) if match.note_table_index == 2]

        self.assertTrue(matches)
        self.assertTrue(all(match.statement_table_index == 1 for match in matches))
        self.assertTrue(all(match.status == "Matched" for match in matches))


class RestatedComparativeTest(unittest.TestCase):
    def test_restated_comparative_note_is_not_reconciled_by_period(self):
        tables = [
            ExtractedTable(
                1,
                [
                    ["CHỈ TIÊU", "Mã số", "Thuyết minh", "Năm nay", "Năm trước"],
                    ["Tổng lợi nhuận kế toán trước thuế", "50", "", "16.544.095.781", "49.127.810.590"],
                ],
                StatementType.INCOME_STATEMENT,
            ),
            _note(
                2,
                [
                    ["", "Trình bày lại trong báo cáo năm nay", "", "Theo báo cáo năm trước"],
                    ["", "VND", "", "VND"],
                    ["Lợi nhuận trước thuế", "1.963.070", "", "2.022.743"],
                ],
                "Thông tin so sánh",
            ),
        ]

        self.assertEqual(
            [match for match in reconcile_note_tables(tables) if match.note_table_index == 2],
            [],
        )


class RuleMessageTest(unittest.TestCase):
    def test_sign_rule_comment_states_expected_sign(self):
        table = ExtractedTable(
            1,
            [
                ["Chỉ tiêu", "Mã số", "Thuyết minh", "Năm nay", "Năm trước"],
                ["Tiền chi trả nợ gốc vay", "34", "", "580.419.000", "(10)"],
            ],
            StatementType.CASH_FLOW,
        )

        _, targets = build_audit_workbook_with_comment_targets([table])

        self.assertIn(
            "Kiểm tra dấu: mã 34 kỳ này phải là số âm (trình bày trong ngoặc), "
            "đang trình bày 580.419.000.",
            [target.message for target in targets],
        )

    def test_retained_earnings_roll_forward_is_review_not_difference(self):
        table = ExtractedTable(
            1,
            [
                ["NGUỒN VỐN", "Mã số", "Thuyết minh", "Số cuối năm", "Số đầu năm"],
                ["Lợi nhuận sau thuế chưa phân phối", "421", "", "10", "(13)"],
                ["- LNST lũy kế đến cuối năm trước", "421a", "", "(14)", "11"],
                ["- LNST năm nay", "421b", "", "24", "(24)"],
            ],
            StatementType.BALANCE_SHEET_EQUITY,
        )

        _, targets = build_audit_workbook_with_comment_targets([table])

        roll_forward = [target for target in targets if "Kiểm tra liên kỳ" in target.message]
        self.assertEqual(len(roll_forward), 1)
        self.assertEqual(roll_forward[0].category, WordCommentCategory.REVIEW)
        self.assertIn("mã 421a kỳ này (14) khác mã 421 kỳ trước trên BS (13)", roll_forward[0].message)


if __name__ == "__main__":
    unittest.main()
