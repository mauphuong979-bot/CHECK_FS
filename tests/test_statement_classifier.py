from __future__ import annotations

import unittest

from src.domain.models import StatementType
from src.extraction.statement_classifier import classify_table, get_title_hint


class StatementClassifierTest(unittest.TestCase):
    def test_get_title_hint_returns_first_non_empty_row(self):
        rows = [
            ["", "   "],
            ["BẢNG CÂN ĐỐI KẾ TOÁN", "Mẫu B 01 - DN"],
        ]
        self.assertEqual(get_title_hint(rows), "BẢNG CÂN ĐỐI KẾ TOÁN | Mẫu B 01 - DN")

    def test_classify_related_party_relationship_table(self):
        rows = [
            ["Bên liên quan", "Mối quan hệ", "Nội dung"],
            ["Công ty A", "Công ty mẹ", "Mua hàng"],
        ]
        statement_type = classify_table(
            rows,
            title_hint="Giao dịch với các bên liên quan",
            context_hints=("Giao dịch với các bên liên quan",),
        )
        self.assertEqual(statement_type, StatementType.NOTE)

    def test_classify_table_of_contents(self):
        rows = [
            ["Nội dung", "Trang"],
            ["1. Báo cáo của Ban Giám đốc", "1"],
            ["2. Báo cáo kiểm toán độc lập", "3"],
        ]
        statement_type = classify_table(
            rows,
            title_hint="MỤC LỤC",
            context_hints=("MụC LỤC",),
        )
        self.assertEqual(statement_type, StatementType.GENERAL_INFO)

    def test_classify_management_information(self):
        rows = [
            ["Họ và tên", "Chức vụ"],
            ["Ông Nguyễn Văn A", "Tổng Giám đốc"],
        ]
        statement_type = classify_table(
            rows,
            title_hint="Danh sách Ban Giám đốc",
            context_hints=("CÁC THÀNH VIÊN BAN GIÁM ĐỐC",),
        )
        self.assertEqual(statement_type, StatementType.GENERAL_INFO)

    def test_classify_business_location_information(self):
        rows = [
            ["Tên chi nhánh", "Địa chỉ"],
            ["Chi nhánh Hà Nội", "123 Phố Huế"],
        ]
        statement_type = classify_table(
            rows,
            title_hint="Địa điểm kinh doanh",
            context_hints=("Địa điểm kinh doanh",),
        )
        self.assertEqual(statement_type, StatementType.GENERAL_INFO)

    def test_classify_signature_block_by_context(self):
        rows = [
            ["Người lập biểu", "Kế toán trưởng", "Tổng Giám đốc"],
            ["(Ký, họ tên)", "(Ký, họ tên)", "(Ký, họ tên, đóng dấu)"],
        ]
        statement_type = classify_table(
            rows,
            title_hint="Thay mặt và đại diện Ban Giám đốc",
            context_hints=("Thay mặt và đại diện",),
        )
        self.assertEqual(statement_type, StatementType.SIGNATURE)

    def test_classify_signature_block_by_text(self):
        rows = [
            ["TUQ. Giám đốc", "", "Kế toán trưởng"],
            ["Nguyễn Văn A", "", "Trần Thị B"],
        ]
        statement_type = classify_table(rows)
        self.assertEqual(statement_type, StatementType.SIGNATURE)

    def test_classify_depreciation_life_range_table(self):
        rows = [
            ["Tài sản", "Năm"],
            ["Nhà cửa, vật kiến trúc", "10 - 20"],
            ["Máy móc, thiết bị", "5 - 10"],
            ["Phương tiện vận tải", "6 - 8"],
        ]
        statement_type = classify_table(
            rows,
            title_hint="Thời gian khấu hao tài sản cố định",
            context_hints=("Khấu hao tài sản",),
        )
        self.assertEqual(statement_type, StatementType.NOTE)

    def test_classify_general_company_info_table(self):
        rows = [
            ["CÔNG TY TNHH ABC"],
            ["Mã số thuế: 0101234567"],
            ["Báo cáo mẫu B 01 - DN"],
        ]
        statement_type = classify_table(rows)
        self.assertEqual(statement_type, StatementType.GENERAL_INFO)

    def test_classify_cash_flow_statement(self):
        rows = [
            ["BÁO CÁO LƯU CHUYỂN TIỀN TỆ", "Mã số", "Thuyết minh", "Năm nay", "Năm trước"],
            ["I. Lưu chuyển tiền từ hoạt động kinh doanh", "01", "", "100.000.000", "80.000.000"],
        ]
        statement_type = classify_table(rows)
        self.assertEqual(statement_type, StatementType.CASH_FLOW)

    def test_classify_balance_sheet_assets(self):
        rows = [
            ["TÀI SẢN", "Mã số", "Thuyết minh", "Số cuối năm", "Số đầu năm"],
            ["A. TÀI SẢN NGẮN HẠN", "100", "", "500.000.000", "400.000.000"],
        ]
        statement_type = classify_table(rows)
        self.assertEqual(statement_type, StatementType.BALANCE_SHEET_ASSETS)

    def test_classify_balance_sheet_equity(self):
        rows = [
            ["NGUỒN VỐN", "Mã số", "Thuyết minh", "Số cuối năm", "Số đầu năm"],
            ["C. NỢ PHẢI TRẢ", "300", "", "200.000.000", "150.000.000"],
        ]
        statement_type = classify_table(rows)
        self.assertEqual(statement_type, StatementType.BALANCE_SHEET_EQUITY)

    def test_classify_income_statement(self):
        rows = [
            ["CHỈ TIÊU", "Mã số", "Thuyết minh", "Năm nay", "Năm trước"],
            ["1. Doanh thu bán hàng và cung cấp dịch vụ", "01", "VI.01", "1.000.000.000", "800.000.000"],
        ]
        statement_type = classify_table(rows)
        self.assertEqual(statement_type, StatementType.INCOME_STATEMENT)

    def test_classify_note_table_by_currency_or_date(self):
        rows = [
            ["Chỉ tiêu", "Số cuối năm", "Số đầu năm"],
            ["Đơn vị tính: VND", "", ""],
            ["Tiền mặt", "50.000.000", "30.000.000"],
        ]
        statement_type = classify_table(rows)
        self.assertEqual(statement_type, StatementType.NOTE)

    def test_classify_note_table_by_short_long_term_labels(self):
        rows = [
            ["Ngắn hạn", "100.000.000"],
            ["Dài hạn", "200.000.000"],
        ]
        statement_type = classify_table(
            rows,
            title_hint="Vay và nợ thuê tài chính",
            context_hints=("Vay và nợ thuê tài chính",),
        )
        self.assertEqual(statement_type, StatementType.NOTE)

    def test_classify_unknown_table(self):
        rows = [
            ["Bảng không rõ nội dung"],
            ["Dữ liệu thử nghiệm"],
        ]
        statement_type = classify_table(rows)
        self.assertEqual(statement_type, StatementType.UNKNOWN)


if __name__ == "__main__":
    unittest.main()
