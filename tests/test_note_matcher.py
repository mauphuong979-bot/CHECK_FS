from decimal import Decimal
import unittest

from src.domain.models import ExtractedTable, StatementType
from src.reconciliation.note_matcher import reconcile_note_tables


def _statement_table(statement_type: StatementType, rows: list[list[str]]) -> ExtractedTable:
    return ExtractedTable(
        index=1,
        rows=[["CHỈ TIÊU", "Mã số", "Thuyết minh", "Năm nay", "Năm trước"], *rows],
        statement_type=statement_type,
    )


def _note_table(
    title: str,
    current: str,
    prior: str = "",
    *,
    context: tuple[str, ...] = (),
) -> ExtractedTable:
    return ExtractedTable(
        index=2,
        rows=[
            ["", "Năm nay", "", "Năm trước"],
            ["", "VND", "", "VND"],
            ["Tổng cộng", current, "", prior],
        ],
        statement_type=StatementType.NOTE,
        title_hint=title,
        context_hints=context,
    )


class NoteMatcherTest(unittest.TestCase):
    def test_parent_detail_is_skipped_when_followed_by_matching_maturity_split(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_ASSETS,
            [
                ["Phải thu ngắn hạn khác", "136", "V.3", "776.000", "-"],
                ["Phải thu dài hạn khác", "216", "V.3", "459.217.279", "459.217.279"],
            ],
        )
        parent = ExtractedTable(
            2,
            [
                ["", "Số cuối năm", "Số đầu năm"],
                ["Ký quỹ thuê văn phòng", "459.217.279", "459.217.279"],
                ["Khác", "776.000", "-"],
                ["Tổng cộng", "459.993.279", "459.217.279"],
            ],
            StatementType.NOTE,
            "3. Phải thu khác",
            context_hints=("3. Phải thu khác",),
        )
        split = ExtractedTable(
            3,
            [
                ["Ngắn hạn", "776.000", "-"],
                ["Dài hạn", "459.217.279", "459.217.279"],
                ["Tổng cộng", "459.993.279", "459.217.279"],
            ],
            StatementType.NOTE,
            "3. Phải thu khác",
            context_hints=("3. Phải thu khác",),
        )

        results = reconcile_note_tables([statement, parent, split])

        self.assertFalse(any(result.note_table_index == 2 for result in results))
        split_results = [result for result in results if result.note_table_index == 3]
        self.assertEqual(len(split_results), 2)
        self.assertEqual({result.status for result in split_results}, {"Matched"})

    def test_informational_depreciation_and_contract_tables_are_not_reconciled(self):
        fixed_assets = ExtractedTable(
            1,
            [
                ["CHỈ TIÊU", "Mã số", "Thuyết minh", "Năm nay", "Năm trước"],
                ["Tài sản cố định hữu hình", "221", "V.8", "30.000", "25.000"],
            ],
            StatementType.BALANCE_SHEET_ASSETS,
        )
        debt = ExtractedTable(
            2,
            [
                ["CHỈ TIÊU", "Mã số", "Thuyết minh", "Năm nay", "Năm trước"],
                ["Vay và nợ thuê tài chính dài hạn", "338", "V.15", "7.011.820.000", "4.030.000.000"],
            ],
            StatementType.BALANCE_SHEET_EQUITY,
        )
        depreciation_life = ExtractedTable(
            3,
            [
                ["", "Năm"],
                ["Nhà cửa, vật kiến trúc", "5 - 7"],
                ["Máy móc, thiết bị", "5 - 10"],
                ["Thiết bị văn phòng", "3 - 7"],
            ],
            StatementType.NOTE,
            "Chính sách khấu hao tài sản cố định",
        )
        fully_depreciated = ExtractedTable(
            4,
            [
                ["", "VND"],
                [
                    "Nguyên giá tài sản cố định hữu hình cuối năm đã khấu hao hết nhưng vẫn còn sử dụng",
                    "26.805.684.650",
                ],
            ],
            StatementType.NOTE,
            "Tài sản cố định hữu hình",
        )
        loan_contracts = ExtractedTable(
            5,
            [
                ["Số hợp đồng", "20221111/HĐVTCN", "202212-001"],
                ["Ngày hợp đồng", "11/11/2022", "10/12/2022"],
                ["Ngày đáo hạn", "11/11/2025", "10/12/2025"],
                ["Hạn mức cho vay (VND)", "5.000.000.000", "4.030.000.000"],
                ["Tài sản đảm bảo", "Không", "Không"],
                ["Số dư cuối kỳ (VND)", "2.981.820.000", "4.030.000.000"],
            ],
            StatementType.NOTE,
            "Vay và nợ thuê tài chính dài hạn",
        )

        results = reconcile_note_tables(
            [fixed_assets, debt, depreciation_life, fully_depreciated, loan_contracts]
        )

        self.assertEqual(results, [])

    def test_column_oriented_contract_schedule_with_total_is_not_reconciled(self):
        debt = ExtractedTable(
            1,
            [
                ["CHỈ TIÊU", "Mã số", "Thuyết minh", "Năm nay", "Năm trước"],
                ["Vay và nợ thuê tài chính dài hạn", "338", "V.15", "567.021.000.000", "400.000.000.000"],
            ],
            StatementType.BALANCE_SHEET_EQUITY,
        )
        contracts = ExtractedTable(
            2,
            [
                ["Số hợp đồng", "Ngày hợp đồng", "Ngày đáo hạn", "Lãi suất", "Số tiền", ""],
                ["", "", "", "(%/năm)", "USD", "VND"],
                ["012023/KINGFA/CN-VN", "22/11/2023", "22/11/2028", "SOFR 12 tháng", "9.000.000", "237.393.000.000"],
                ["022023/KINGFA/CN-VN", "01/12/2023", "01/12/2028", "SOFR 12 tháng", "5.000.000", "131.885.000.000"],
                ["Tổng cộng", "", "", "", "26.000.000", "685.802.000.000"],
            ],
            StatementType.NOTE,
            "Vay và nợ thuê tài chính dài hạn",
        )

        self.assertEqual(reconcile_note_tables([debt, contracts]), [])

    def test_tt200_equity_movement_reconciles_411_421_and_pl_60_by_period(self):
        balance_sheet = ExtractedTable(
            index=1,
            rows=[
                ["CHỈ TIÊU", "Mã số", "Thuyết minh", "Năm nay", "Năm trước"],
                ["Vốn góp của chủ sở hữu", "411", "V.16", "7.294.502.372", "7.294.502.372"],
                ["Lợi nhuận sau thuế chưa phân phối", "421", "V.16", "(8.700.547.294)", "(28.155.585.793)"],
            ],
            statement_type=StatementType.BALANCE_SHEET_EQUITY,
        )
        income_statement = ExtractedTable(
            index=2,
            rows=[
                ["CHỈ TIÊU", "Mã số", "Thuyết minh", "Năm nay", "Năm trước"],
                ["Lợi nhuận sau thuế thu nhập doanh nghiệp", "60", "", "19.455.038.499", "9.896.243.076"],
            ],
            statement_type=StatementType.INCOME_STATEMENT,
        )
        note = ExtractedTable(
            index=3,
            rows=[
                ["Tình hình tăng giảm vốn chủ sở hữu", "Vốn góp của chủ sở hữu", "", "Lợi nhuận sau thuế chưa phân phối", "", "Tổng cộng"],
                ["", "VND", "", "VND", "", "VND"],
                ["Số đầu năm trước", "7.294.502.372", "", "(38.051.828.869)", "", "(30.757.326.497)"],
                ["Lợi nhuận thuần trong năm trước", "-", "", "9.896.243.076", "", "9.896.243.076"],
                ["Số cuối năm trước, đầu năm nay", "7.294.502.372", "", "(28.155.585.793)", "", "(20.861.083.421)"],
                ["Lợi nhuận thuần trong năm", "-", "", "19.455.038.499", "", "19.455.038.499"],
                ["Số cuối năm", "7.294.502.372", "", "(8.700.547.294)", "", "(1.406.044.922)"],
            ],
            statement_type=StatementType.NOTE,
            title_hint="16. Vốn chủ sở hữu",
        )

        results = reconcile_note_tables([balance_sheet, income_statement, note])
        by_code = {result.statement_code: result for result in results}

        self.assertEqual(set(by_code), {"411", "421", "60"})
        self.assertTrue(all(result.status == "Matched" for result in results))
        self.assertEqual(by_code["411"].note_current_cell, "B10")
        self.assertEqual(by_code["411"].note_prior_cell, "B8")
        self.assertEqual(by_code["421"].note_current, Decimal("-8700547294"))
        self.assertEqual(by_code["421"].note_prior, Decimal("-28155585793"))
        self.assertEqual(by_code["60"].note_current_cell, "D9")
        self.assertEqual(by_code["60"].note_prior_cell, "D7")

    def test_tt200_equity_movement_reports_only_missing_target_item(self):
        balance_sheet = _statement_table(
            StatementType.BALANCE_SHEET_EQUITY,
            [["Vốn góp của chủ sở hữu", "411", "V.16", "100", "90"]],
        )
        note = ExtractedTable(
            index=2,
            rows=[
                ["", "Vốn góp của chủ sở hữu", "Lợi nhuận sau thuế chưa phân phối"],
                ["Số cuối năm trước, đầu năm nay", "90", "40"],
                ["Số cuối năm", "100", "50"],
            ],
            statement_type=StatementType.NOTE,
            title_hint="Vốn chủ sở hữu",
        )

        results = reconcile_note_tables([balance_sheet, note])
        by_code = {result.statement_code: result for result in results}

        self.assertEqual(by_code["411"].status, "Matched")
        self.assertEqual(by_code["421"].status, "Statement not found")
        self.assertNotIn("60", by_code)

    def test_receivable_net_continuation_is_not_reconciled_as_independent_bs_value(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_ASSETS,
            [["Phải thu ngắn hạn của khách hàng", "131", "V.2", "100", "50"]],
        )
        detail = ExtractedTable(
            2,
            [["Tổng cộng", "100", "50"]],
            StatementType.NOTE,
            "Phải thu ngắn hạn của khách hàng",
        )
        continuation = ExtractedTable(
            3,
            [["Dự phòng phải thu ngắn hạn khó đòi", "(10)", "(5)"], ["Giá trị thuần", "90", "45"]],
            StatementType.NOTE,
            "Phải thu ngắn hạn của khách hàng",
        )

        results = reconcile_note_tables([statement, detail, continuation])

        self.assertFalse(any(result.note_table_index == 3 for result in results))

    def test_income_tax_schedule_reconciles_profit_and_current_tax_expense_to_pl(self):
        statement = _statement_table(
            StatementType.INCOME_STATEMENT,
            [
                ["14. Tổng lợi nhuận kế toán trước thuế", "50", "", "(10.662)", "(180.810)"],
                ["15. Chi phí thuế thu nhập doanh nghiệp hiện hành", "51", "VI.9", "-", "-"],
                ["11. Thu nhập khác", "31", "VI.7", "5.000", "1.000"],
            ],
        )
        note = ExtractedTable(
            index=2,
            rows=[
                ["", "Năm nay", "", "Năm trước"],
                ["", "VND", "", "VND"],
                ["Lợi nhuận trước thuế", "(10.662)", "", "(180.810)"],
                ["Thu nhập chịu thuế", "20", "", "10"],
                ["Chi phí thuế thu nhập doanh nghiệp hiện hành", "-", "", "-"],
            ],
            statement_type=StatementType.NOTE,
            title_hint="Đối với các khoản thu nhập khác không được ưu đãi",
            context_hints=("Chi phí thuế thu nhập doanh nghiệp hiện hành",),
        )

        results = reconcile_note_tables([statement, note])

        self.assertEqual({result.statement_code for result in results}, {"50", "51"})
        self.assertTrue(all(result.status == "Matched" for result in results))
        self.assertTrue(all(result.match_type == "Tên dòng + PL" for result in results))

    def test_income_tax_schedule_maps_date_range_headers_to_distinct_periods(self):
        statement = _statement_table(
            StatementType.INCOME_STATEMENT,
            [["14. Tổng lợi nhuận kế toán trước thuế", "50", "", "(6.139.421.945)", "(3.305.515.469)"]],
        )
        note = ExtractedTable(
            index=2,
            rows=[
                ["", "Từ 01/01/2025 đến 31/12/2025", "", "Từ 14/05/2024 đến 31/12/2024"],
                ["", "VND", "", "VND"],
                ["Lợi nhuận trước thuế", "(6.139.421.945)", "", "(3.305.515.469)"],
            ],
            statement_type=StatementType.NOTE,
        )

        result = reconcile_note_tables([statement, note])[0]

        self.assertEqual(result.note_current_cell, "B6")
        self.assertEqual(result.note_prior_cell, "D6")
        self.assertEqual(result.status, "Matched")

    def test_income_tax_schedule_selects_matching_currency_when_two_pl_tables_exist(self):
        usd_statement = _statement_table(
            StatementType.INCOME_STATEMENT,
            [
                ["Tổng lợi nhuận kế toán trước thuế", "50", "", "636.654,19", "1.963.070,83"],
                ["Chi phí thuế thu nhập doanh nghiệp hiện hành", "51", "VI.8", "120.503,26", "362.531,30"],
            ],
        )
        vnd_statement = ExtractedTable(
            index=3,
            rows=[
                ["CHỈ TIÊU", "Mã số", "Thuyết minh", "Năm nay", "Năm trước"],
                ["Tổng lợi nhuận kế toán trước thuế", "50", "", "16.544.095.781", "49.127.810.590"],
                ["Chi phí thuế thu nhập doanh nghiệp hiện hành", "51", "VI.8", "3.131.397.714", "9.072.708.314"],
            ],
            statement_type=StatementType.INCOME_STATEMENT,
        )
        note = ExtractedTable(
            index=2,
            rows=[
                ["", "Năm nay", "", "Năm trước"],
                ["", "USD", "", "USD"],
                ["Lợi nhuận trước thuế", "636.654,19", "", "1.963.070,83"],
                ["Chi phí thuế thu nhập doanh nghiệp hiện hành", "120.503,26", "", "362.531,30"],
            ],
            statement_type=StatementType.NOTE,
        )

        results = reconcile_note_tables([usd_statement, note, vnd_statement])

        self.assertEqual({result.statement_table_index for result in results}, {1})
        self.assertTrue(all(result.status == "Matched" for result in results))

    def test_maturity_rows_match_distinct_balance_sheet_items(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_ASSETS,
            [
                ["Phải thu ngắn hạn khác", "136", "V.5", "249.350.916", "1.183.577.037"],
                ["Phải thu dài hạn khác", "216", "V.5", "275.889.600", "-"],
            ],
        )
        note = ExtractedTable(
            index=2,
            rows=[
                ["Ngắn hạn", "249.350.916", "1.183.577.037"],
                ["Dài hạn", "275.889.600", "-"],
                ["Tổng cộng", "525.240.516", "1.183.577.037"],
            ],
            statement_type=StatementType.NOTE,
            title_hint="5. Phải thu khác",
            context_hints=("5. Phải thu khác",),
        )

        results = reconcile_note_tables([statement, note])

        self.assertEqual({result.statement_code for result in results}, {"136", "216"})
        self.assertTrue(all(result.status == "Matched" for result in results))
        by_code = {result.statement_code: result for result in results}
        self.assertEqual(by_code["136"].note_current_cell, "B4")
        self.assertEqual(by_code["136"].note_prior_cell, "C4")
        self.assertEqual(by_code["216"].note_current_cell, "B5")
        self.assertEqual(by_code["216"].note_prior, Decimal(0))

    def test_combined_maturity_title_matches_short_and_long_term_bs_items(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_ASSETS,
            [
                ["Phải thu ngắn hạn khác", "136", "V.4", "27.469.821", "17.078.778"],
                ["Phải thu dài hạn khác", "216", "V.4", "317.520.000", "317.520.000"],
            ],
        )
        note = ExtractedTable(
            index=2,
            rows=[
                ["Ngắn hạn", "27.469.821", "", "17.078.778"],
                ["Dài hạn", "317.520.000", "", "317.520.000"],
                ["Tổng cộng", "344.989.821", "", "334.598.778"],
            ],
            statement_type=StatementType.NOTE,
            title_hint="4. Phải thu ngắn hạn/dài hạn khác",
            context_hints=("Phải thu ngắn hạn/dài hạn khác",),
        )

        results = reconcile_note_tables([statement, note])

        self.assertEqual({result.statement_code for result in results}, {"136", "216"})
        self.assertTrue(all(result.status == "Matched" for result in results))
        self.assertTrue(all(result.note_code == "V.4" for result in results))

    def test_state_tax_balance_rows_reconcile_all_four_values_to_bs(self):
        assets = _statement_table(
            StatementType.BALANCE_SHEET_ASSETS,
            [["Thuế và các khoản khác phải thu Nhà nước", "153", "V.14", "100.097.721", "110.855.895"]],
        )
        equity = ExtractedTable(
            index=3,
            rows=[
                ["CHỈ TIÊU", "Mã số", "Thuyết minh", "Năm nay", "Năm trước"],
                ["Thuế và các khoản phải nộp Nhà nước", "313", "V.14", "-", "-"],
            ],
            statement_type=StatementType.BALANCE_SHEET_EQUITY,
        )
        note = ExtractedTable(
            index=2,
            rows=[
                ["", "Số cuối năm", "", "Số đầu năm"],
                ["", "VND", "", "VND"],
                ["Thuế và các khoản khác phải thu Nhà nước", "100.097.721", "", "110.855.895"],
                ["Thuế và các khoản phải nộp Nhà nước", "-", "", "-"],
            ],
            statement_type=StatementType.NOTE,
            title_hint="14. Thuế và các khoản phải nộp Nhà nước",
        )

        results = reconcile_note_tables([assets, note, equity])

        self.assertEqual({result.statement_code for result in results}, {"153", "313"})
        self.assertTrue(all(result.status == "Matched" for result in results))
        self.assertTrue(all(result.match_type == "Tên dòng + mã BS" for result in results))
        by_code = {result.statement_code: result for result in results}
        self.assertEqual(by_code["153"].note_current, Decimal("100097721"))
        self.assertEqual(by_code["153"].note_prior, Decimal("110855895"))
        self.assertEqual(by_code["313"].note_current, Decimal(0))
        self.assertEqual(by_code["313"].note_prior, Decimal(0))

    def test_contributed_capital_without_total_row_reconciles_to_bs_code_411(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_EQUITY,
            [
                ["Vốn chủ sở hữu", "410", "V.14", "35.640", "13.866"],
                ["Vốn góp của chủ sở hữu", "411", "", "10.000", "10.000"],
            ],
        )
        note = ExtractedTable(
            index=2,
            rows=[
                ["Theo giấy chứng nhận đăng ký doanh nghiệp", "", "", "Vốn đã góp", "", "Vốn đã góp"],
                ["", "", "", "Số cuối năm", "", "Số đầu năm"],
                ["", "VND", "", "VND", "", "VND"],
                ["Cổ đông A", "6.000", "", "6.000", "", "6.000"],
                ["Cổ đông B", "4.000", "", "4.000", "", "4.000"],
            ],
            statement_type=StatementType.NOTE,
            title_hint="Tình hình góp vốn điều lệ của Công ty",
            context_hints=("14. Vốn chủ sở hữu",),
        )

        result = reconcile_note_tables([statement, note])[0]

        self.assertEqual(result.statement_code, "411")
        self.assertEqual(result.note_code, "V.14")
        self.assertEqual(result.note_current, Decimal("10000"))
        self.assertEqual(result.note_prior, Decimal("10000"))
        self.assertEqual(result.status, "Matched")
        self.assertEqual(result.match_type, "Vốn đã góp + mã BS")

    def test_contributed_capital_with_total_uses_vnd_not_usd(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_EQUITY,
            [["Vốn góp của chủ sở hữu", "411", "V.14", "665.183.950.966", "665.183.950.966"]],
        )
        note = ExtractedTable(
            2,
            [
                ["Theo giấy chứng nhận đăng ký đầu tư", "", "", "", "", "Vốn đã góp", "", ""],
                ["chứng nhận đăng ký doanh nghiệp", "", "", "Số cuối năm", "", "", "", "Số đầu năm"],
                ["", "USD", "", "USD", "", "VND", "", "VND"],
                ["Kingfa Sci. And Tech. Co., Ltd.", "19.600.000", "", "19.600.000", "", "466.970.000.000", "", "466.970.000.000"],
                ["Hongkong Kingfa Development Co., Limited", "8.400.000", "", "8.400.000", "", "198.213.950.966", "", "198.213.950.966"],
                ["Tổng cộng", "28.000.000", "", "28.000.000", "", "665.183.950.966", "", "665.183.950.966"],
            ],
            StatementType.NOTE,
            "Vốn góp của chủ sở hữu",
        )

        result = reconcile_note_tables([statement, note])[0]

        self.assertEqual(result.statement_code, "411")
        self.assertEqual(result.note_current, Decimal("665183950966"))
        self.assertEqual(result.note_prior, Decimal("665183950966"))
        self.assertEqual(result.note_current_cell, "F9")
        self.assertEqual(result.note_prior_cell, "H9")
        self.assertEqual(result.status, "Matched")

    def test_equity_movement_is_recognized_from_table_heading_when_title_is_narrow(self):
        balance_sheet = _statement_table(
            StatementType.BALANCE_SHEET_EQUITY,
            [
                ["Vốn góp của chủ sở hữu", "411", "V.16", "665.183.950.966", "665.183.950.966"],
                ["Lợi nhuận sau thuế chưa phân phối", "421", "V.16", "229.680.337.399", "(12.681.636.187)"],
            ],
        )
        income_statement = _statement_table(
            StatementType.INCOME_STATEMENT,
            [["Lợi nhuận sau thuế thu nhập doanh nghiệp", "60", "", "242.361.973.586", "(14.665.951.431)"]],
        )
        note = ExtractedTable(
            3,
            [
                ["Tình hình tăng giảm vốn chủ sở hữu", "Vốn góp của chủ sở hữu", "", "Lợi nhuận sau thuế chưa phân phối", "", "Tổng cộng"],
                ["", "VND", "", "VND", "", "VND"],
                ["", "", "", "", "", ""],
                ["Số đầu năm trước", "665.183.950.966", "", "1.984.315.244", "", "667.168.266.210"],
                ["Lợi nhuận thuần trong năm trước", "-", "", "(14.665.951.431)", "", "(14.665.951.431)"],
                ["Số cuối năm trước, đầu năm nay", "665.183.950.966", "", "(12.681.636.187)", "", "652.502.314.779"],
                ["Lợi nhuận thuần trong năm", "-", "", "242.361.973.586", "", "242.361.973.586"],
                ["Số cuối năm", "665.183.950.966", "", "229.680.337.399", "", "894.864.288.365"],
            ],
            StatementType.NOTE,
            "Vốn góp của chủ sở hữu",
        )

        results = reconcile_note_tables([balance_sheet, income_statement, note])
        by_code = {result.statement_code: result for result in results}

        self.assertEqual(set(by_code), {"411", "421", "60"})
        self.assertTrue(all(result.status == "Matched" for result in results))
        self.assertEqual(by_code["411"].note_current_cell, "B11")
        self.assertEqual(by_code["421"].note_prior_cell, "D9")
        self.assertEqual(by_code["60"].note_current_cell, "D10")

    def test_informational_notes_are_excluded_from_reconciliation(self):
        tables = [
            ExtractedTable(
                1,
                [["", "Số năm khấu hao"], ["Nhà cửa", "11 - 35"]],
                StatementType.NOTE,
                "Thời gian khấu hao",
            ),
            ExtractedTable(
                2,
                [["", "VND"], ["Giá trị tài sản dùng để thế chấp", "222.547.876.933"]],
                StatementType.NOTE,
                "Tài sản thế chấp",
            ),
            ExtractedTable(
                3,
                [["", "Số cuối năm", "Số đầu năm"], ["Ngoại tệ", "891", "808"]],
                StatementType.NOTE,
                "Các khoản mục ngoài Bảng cân đối kế toán",
            ),
        ]

        self.assertEqual(reconcile_note_tables(tables), [])

    def test_related_party_and_future_lease_notes_are_excluded_from_reconciliation(self):
        tables = [
            ExtractedTable(
                1,
                [
                    ["", "Năm nay", "Năm trước"],
                    ["Mượn tiền", "", ""],
                    ["Ông A", "100", "90"],
                ],
                StatementType.NOTE,
                "Các nghiệp vụ kinh tế quan trọng với các bên liên quan được trình bày ở bảng sau",
                context_hints=("Giao dịch với các bên liên quan",),
            ),
            ExtractedTable(
                2,
                [["Bên liên quan", "Mối quan hệ"], ["Ông A", "Giám đốc"]],
                StatementType.NOTE,
                "Giao dịch với các bên liên quan",
            ),
            ExtractedTable(
                3,
                [
                    ["", "Năm nay", "Năm trước"],
                    ["Trả trong vòng một năm", "40", "30"],
                    ["Sau một năm", "60", "50"],
                    ["Tổng cộng", "100", "80"],
                ],
                StatementType.NOTE,
                "Các khoản tiền thuê phải trả trong tương lai theo các hợp đồng thuê",
            ),
        ]

        self.assertEqual(reconcile_note_tables(tables), [])

    def test_asset_notes_reconcile_three_blocks_by_statement_codes(self):
        cases = (
            ("Tài sản cố định hữu hình", ("221", "222", "223")),
            ("Tài sản cố định thuê tài chính", ("224", "225", "226")),
            ("Tài sản cố định vô hình", ("227", "228", "229")),
            ("Bất động sản đầu tư", ("230", "231", "232")),
        )
        for title, (net_code, cost_code, depreciation_code) in cases:
            with self.subTest(title=title):
                depreciation_section_label = (
                    "Giá trị khấu hao"
                    if title == "Tài sản cố định vô hình"
                    else "Giá trị hao mòn"
                )
                statement = _statement_table(
                    StatementType.BALANCE_SHEET_ASSETS,
                    [
                        ["Giá trị còn lại", net_code, "V.1", "80", "70"],
                        ["Nguyên giá", cost_code, "", "100", "90"],
                        ["Giá trị hao mòn lũy kế", depreciation_code, "", "(20)", "(20)"],
                    ],
                )
                note = ExtractedTable(
                    index=2,
                    rows=[
                        ["", "Chi tiết", "Tổng cộng"],
                        ["Nguyên giá", "", ""],
                        ["Số đầu năm", "", "90"],
                        ["Số cuối năm", "", "100"],
                        [depreciation_section_label, "", ""],
                        ["Số đầu năm", "", "(20)"],
                        ["Số cuối năm", "", "(20)"],
                        ["Giá trị còn lại", "", ""],
                        ["Số đầu năm", "", "70"],
                        ["Số cuối năm", "", "80"],
                    ],
                    statement_type=StatementType.NOTE,
                    title_hint=title,
                )

                results = reconcile_note_tables([statement, note])

                self.assertEqual(
                    {result.statement_code for result in results},
                    {net_code, cost_code, depreciation_code},
                )
                self.assertTrue(all(result.status == "Matched" for result in results))
                self.assertTrue(all(result.match_type == "Mã số BCTC" for result in results))
                by_code = {result.statement_code: result for result in results}
                self.assertEqual(by_code[cost_code].note_current_cell, "C7")
                self.assertEqual(by_code[cost_code].note_prior_cell, "C6")
                self.assertEqual(by_code[depreciation_code].note_current_cell, "C10")
                self.assertEqual(by_code[depreciation_code].note_prior_cell, "C9")
                self.assertEqual(by_code[net_code].note_current_cell, "C13")
                self.assertEqual(by_code[net_code].note_prior_cell, "C12")

    def test_compact_intangible_asset_rows_reconcile_all_three_bs_codes(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_ASSETS,
            [
                ["Giá trị còn lại", "227", "V.8", "280.000.004", "-"],
                ["Nguyên giá", "228", "", "350.000.000", "-"],
                ["Giá trị hao mòn lũy kế", "229", "", "(69.999.996)", "-"],
            ],
        )
        note = ExtractedTable(
            2,
            [
                ["Phần mềm CDMS", "Số cuối năm", "", "Số đầu năm"],
                ["", "VND", "", "VND"],
                ["", "", "", ""],
                ["Nguyên giá", "350.000.000", "", "-"],
                ["Khấu hao lũy kế", "(69.999.996)", "", "-"],
                ["Giá trị còn lại", "280.000.004", "", "-"],
                ["Khấu hao trong năm", "(69.999.996)", "", "-"],
            ],
            StatementType.NOTE,
            "Tài sản cố định vô hình",
        )

        results = reconcile_note_tables([statement, note])
        by_code = {result.statement_code: result for result in results}

        self.assertEqual(set(by_code), {"227", "228", "229"})
        self.assertTrue(all(result.status == "Matched" for result in results))
        self.assertEqual(by_code["228"].note_current_cell, "B7")
        self.assertEqual(by_code["229"].note_current_cell, "B8")
        self.assertEqual(by_code["227"].note_current_cell, "B9")
        self.assertTrue(all(result.note_prior == Decimal(0) for result in results))

    def test_asset_note_reports_missing_statement_code_without_value_fallback(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_ASSETS,
            [["Giá trị còn lại", "221", "V.1", "80", "70"]],
        )
        note = ExtractedTable(
            index=2,
            rows=[
                ["", "Tổng cộng"],
                ["Nguyên giá", ""],
                ["Số đầu năm", "90"],
                ["Số cuối năm", "100"],
                ["Giá trị hao mòn", ""],
                ["Số đầu năm", "(20)"],
                ["Số cuối năm", "(20)"],
                ["Giá trị còn lại", ""],
                ["Số đầu năm", "70"],
                ["Số cuối năm", "80"],
            ],
            statement_type=StatementType.NOTE,
            title_hint="Tài sản cố định hữu hình",
        )

        results = reconcile_note_tables([statement, note])
        status_by_code = {result.statement_code: result.status for result in results}

        self.assertEqual(status_by_code["221"], "Matched")
        self.assertEqual(status_by_code["222"], "Statement not found")
        self.assertEqual(status_by_code["223"], "Statement not found")
        self.assertTrue(all(result.match_type == "Mã số BCTC" for result in results))

    def test_exact_balance_sheet_match(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_ASSETS,
            [["Hàng tồn kho", "140", "V.6", "1.000", "900"]],
        )

        result = reconcile_note_tables([statement, _note_table("Hàng tồn kho", "1.000", "900")])

        self.assertEqual(result[0].source, "BS")
        self.assertEqual(result[0].statement_code, "140")
        self.assertEqual(result[0].status, "Matched")

    def test_exact_income_statement_match(self):
        statement = _statement_table(
            StatementType.INCOME_STATEMENT,
            [["1. Doanh thu thuần", "10", "VI.1", "2.000", "1.500"]],
        )

        result = reconcile_note_tables([statement, _note_table("Doanh thu thuần", "2.000", "1.500")])

        self.assertEqual(result[0].source, "PL")
        self.assertEqual(result[0].note_code, "VI.1")
        self.assertEqual(result[0].status, "Matched")

    def test_exact_cash_flow_match(self):
        statement = _statement_table(
            StatementType.CASH_FLOW,
            [["Tiền cuối kỳ", "70", "VII.1", "(2.000)", "1.000"]],
        )

        result = reconcile_note_tables([statement, _note_table("Tiền cuối kỳ", "(2.000)", "1.000")])

        self.assertEqual(result[0].source, "CF")
        self.assertEqual(result[0].statement_current, Decimal("-2000"))
        self.assertEqual(result[0].status, "Matched")

    def test_depreciation_expense_reconciles_to_cash_flow_code_02_by_absolute_value(self):
        statement = _statement_table(
            StatementType.CASH_FLOW,
            [[
                "Khấu hao tài sản cố định và bất động sản đầu tư",
                "02",
                "",
                "(37.776.600.586)",
                "36.138.266.974",
            ]],
        )
        note = ExtractedTable(
            index=2,
            rows=[
                ["", "Năm nay", "Năm trước"],
                ["", "VND", "VND"],
                ["Chi phí nguyên liệu, vật liệu", "10", "9"],
                ["Chi phí khấu hao tài sản cố định", "37.776.600.586", "36.138.266.974"],
                ["Tổng cộng", "37.776.600.596", "36.138.266.983"],
            ],
            statement_type=StatementType.NOTE,
            title_hint="Chi phí sản xuất, kinh doanh theo yếu tố",
        )

        result = reconcile_note_tables([statement, note])[0]

        self.assertEqual(result.source, "CF")
        self.assertEqual(result.statement_code, "02")
        self.assertEqual(result.match_type, "Tên dòng + mã CF")
        self.assertEqual(result.statement_current, Decimal("-37776600586"))
        self.assertEqual(result.note_current, Decimal("37776600586"))
        self.assertEqual(result.current_difference, Decimal(0))
        self.assertEqual(result.status, "Matched")

    def test_not_found(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_ASSETS,
            [["Hàng tồn kho", "140", "V.6", "1.000", "900"]],
        )

        result = reconcile_note_tables([statement, _note_table("Nội dung không liên quan", "15")])

        self.assertEqual(len(result), 1)
        self.assertEqual(result[0].status, "Statement not found")
        self.assertEqual(result[0].match_type, "Tên bảng + Tổng cộng")

    def test_tax_rollforward_is_not_reconciled_to_a_single_statement_item(self):
        table = ExtractedTable(
            1,
            [
                ["", "Số đầu năm", "Số phải nộp", "Số đã nộp, khấu trừ", "Số cuối năm"],
                ["", "VND", "VND", "VND", "VND"],
                ["Thuế giá trị gia tăng", "-", "100", "(100)", "-"],
                ["Thuế thu nhập cá nhân", "20", "30", "(10)", "40"],
                ["Tổng cộng", "20", "130", "(110)", "40"],
            ],
            StatementType.NOTE,
            "Thuế và các khoản phải nộp Nhà nước",
        )

        self.assertEqual(reconcile_note_tables([table]), [])

    def test_each_item_with_same_note_reference_has_independent_status(self):
        statement = _statement_table(
            StatementType.INCOME_STATEMENT,
            [
                ["Doanh thu bán hàng", "01", "VI.1", "2.100", "1.600"],
                ["Doanh thu thuần", "10", "VI.1", "2.000", "1.500"],
            ],
        )

        results = reconcile_note_tables([statement, _note_table("Doanh thu thuần", "2.000", "1.500")])

        self.assertEqual({result.statement_code for result in results}, {"01", "10"})
        status_by_code = {result.statement_code: result.status for result in results}
        self.assertEqual(status_by_code["10"], "Matched")
        self.assertEqual(status_by_code["01"], "Note value not found")
        self.assertTrue(all(result.match_type == "Tham chiếu Thuyết minh" for result in results))

    def test_slightly_different_name_without_reference_is_not_inferred(self):
        statement = _statement_table(
            StatementType.INCOME_STATEMENT,
            [["Chi phí quản lý doanh nghiệp", "26", "", "1.000", "900"]],
        )

        result = reconcile_note_tables(
            [statement, _note_table("Chi phí quản lý DN", "1.000", "900")]
        )

        self.assertEqual(result[0].match_type, "Tên bảng + Tổng cộng")
        self.assertEqual(result[0].status, "Statement not found")
        self.assertEqual(result[0].statement_code, "")

    def test_difference_of_one_is_reported_with_two_decimal_rounding(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_ASSETS,
            [["Tiền", "110", "V.1", "1.001", "900"]],
        )

        result = reconcile_note_tables([statement, _note_table("Tiền", "1.000", "900")])

        self.assertEqual(result[0].current_difference, Decimal("1"))
        self.assertEqual(result[0].status, "Difference")

    def test_material_difference_is_reported(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_ASSETS,
            [["Tiền", "110", "V.1", "1.010", "900"]],
        )

        result = reconcile_note_tables([statement, _note_table("Tiền", "1.000", "900")])

        self.assertEqual(result[0].status, "Difference")

    def test_parent_note_context_can_link_multiple_statement_items(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_EQUITY,
            [
                ["Vay và nợ thuê tài chính ngắn hạn", "320", "V.15", "1.000", "900"],
                ["Vay và nợ thuê tài chính dài hạn", "338", "V.15", "500", "400"],
            ],
        )
        note = _note_table("Ngắn hạn", "1.000", "900", context=("Vay và nợ thuê tài chính",))

        results = reconcile_note_tables([statement, note])

        self.assertEqual({result.statement_code for result in results}, {"320"})
        self.assertTrue(all(result.note_code == "V.15" for result in results))
        self.assertEqual(results[0].status, "Matched")
        self.assertEqual(results[0].match_type, "Tiêu đề kỳ hạn và Thuyết minh")

    def test_maturity_subtitle_reports_expected_item_when_bs_target_is_missing(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_EQUITY,
            [["Phải trả người bán", "311", "V.14", "100", "90"]],
        )
        note = _note_table(
            "15.1. Ngắn hạn",
            "1.000",
            "900",
            context=("15. Vay và nợ thuê tài chính",),
        )

        results = reconcile_note_tables([statement, note])

        self.assertEqual(results[0].status, "Statement not found")
        self.assertEqual(results[0].source, "BS")
        self.assertEqual(results[0].item_name, "Vay và nợ thuê tài chính ngắn hạn")
        self.assertEqual(results[0].note_current_cell, "B6")

    def test_maturity_subtitle_reports_difference_instead_of_missing_note_value(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_EQUITY,
            [["Vay và nợ thuê tài chính ngắn hạn", "320", "V.15", "1.100", "900"]],
        )
        note = _note_table(
            "15.1. Ngắn hạn",
            "1.000",
            "900",
            context=("15. Vay và nợ thuê tài chính",),
        )

        result = reconcile_note_tables([statement, note])[0]

        self.assertEqual(result.status, "Difference")
        self.assertEqual(result.current_difference, Decimal("100"))
        self.assertEqual(result.note_current_cell, "B6")

    def test_full_item_title_containing_maturity_is_not_treated_as_subtitle(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_ASSETS,
            [["Phải thu ngắn hạn của khách hàng", "131", "V.2", "1.000", "900"]],
        )
        note = _note_table("2. Phải thu ngắn hạn của khách hàng", "1.000", "900")

        result = reconcile_note_tables([statement, note])[0]

        self.assertEqual(result.status, "Matched")
        self.assertEqual(result.statement_code, "131")
        self.assertEqual(result.match_type, "Tham chiếu Thuyết minh")

    def test_equity_note_uses_vnd_total_pair_to_confirm_unique_bs_item(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_EQUITY,
            [
                ["Vốn chủ sở hữu", "400", "V.16", "2.411.812.417.162", "880.676.943.165"],
                ["Vốn góp của chủ sở hữu", "411", "", "887.657.417.162", "529.677.417.162"],
                ["Thặng dư vốn cổ phần", "412", "V.16", "20", "10"],
            ],
        )
        note = ExtractedTable(
            index=2,
            rows=[
                ["", "", "Vốn đã góp", ""],
                ["", "Số cuối năm", "Số cuối năm", "Số đầu năm"],
                ["", "USD", "VND", "VND"],
                ["Thrive Nation Group Limited", "36.500.000", "887.657.417.162", "529.677.417.162"],
                ["Tổng cộng", "36.500.000", "887.657.417.162", "529.677.417.162"],
            ],
            statement_type=StatementType.NOTE,
            title_hint="16. Vốn chủ sở hữu",
        )

        result = reconcile_note_tables([statement, note])[0]

        self.assertEqual(result.statement_code, "411")
        self.assertEqual(result.status, "Matched")
        self.assertEqual(result.match_type, "Vốn đã góp + mã BS")
        self.assertEqual(result.note_code, "V.16")
        self.assertEqual(result.note_current, Decimal("887657417162"))
        self.assertEqual(result.note_prior, Decimal("529677417162"))
        self.assertEqual(len(reconcile_note_tables([statement, note])), 1)

    def test_equity_movement_table_with_so_cuoi_nam_nay(self):
        statement = _statement_table(
            StatementType.BALANCE_SHEET_EQUITY,
            [
                ["VỐN CHỦ SỞ HỮU", "400", "V.13", "27.506.306.836", "4.186.970.084"],
                ["Vốn góp của chủ sở hữu", "411", "", "17.263.927.091", "17.263.927.091"],
                ["Lợi nhuận sau thuế chưa phân phối", "421", "", "10.242.379.745", "-13.076.957.007"],
            ],
        )
        note = ExtractedTable(
            index=2,
            rows=[
                ["Tình hình tăng, giảm vốn chủ sở hữu", "Vốn góp của chủ sở hữu", "Lợi nhuận sau thuế chưa phân phối", "Tổng cộng"],
                ["", "VND", "VND", "VND"],
                ["Số đầu năm trước", "17.263.927.091", "11.360.666.756", "28.624.593.847"],
                ["Lợi nhuận thuần trong năm trước", "-", "(24.437.623.763)", "(24.437.623.763)"],
                ["Số cuối năm trước, đầu năm nay", "17.263.927.091", "(13.076.957.007)", "4.186.970.084"],
                ["Lợi nhuận thuần trong năm nay", "-", "24.066.148.360", "24.066.148.360"],
                ["Số cuối năm nay", "17.263.927.091", "10.242.379.745", "27.506.306.836"],
            ],
            statement_type=StatementType.NOTE,
            title_hint="Vốn chủ sở hữu",
        )

        results = reconcile_note_tables([statement, note])
        matched_codes = {r.statement_code: r.status for r in results}
        self.assertEqual(matched_codes.get("411"), "Matched")
        self.assertEqual(matched_codes.get("421"), "Matched")
        self.assertNotIn("400", matched_codes)


if __name__ == "__main__":
    unittest.main()

