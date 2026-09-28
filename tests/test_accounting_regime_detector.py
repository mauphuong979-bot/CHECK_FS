import unittest

from src.domain.accounting_regime_detector import detect_accounting_regime
from src.domain.models import AccountingRegime, DetectionConfidence, ExtractedTable, StatementType


class AccountingRegimeDetectorTest(unittest.TestCase):
    def test_detects_explicit_tt200_reference(self):
        detection = detect_accounting_regime([
            ExtractedTable(
                1,
                [["Chế độ kế toán áp dụng theo Thông tư số 200/2014/TT-BTC"]],
                StatementType.GENERAL_INFO,
            )
        ])

        self.assertEqual(detection.regime, AccountingRegime.TT200_2014)
        self.assertEqual(detection.confidence, DetectionConfidence.HIGH)

    def test_detects_tt133_and_tt99_references(self):
        for text, expected in (
            ("Áp dụng TT 133/2016/TT-BTC", AccountingRegime.TT133_2016),
            ("Theo Thông tư 99/2025/TT-BTC", AccountingRegime.TT99_2025),
        ):
            with self.subTest(text=text):
                detection = detect_accounting_regime([
                    ExtractedTable(1, [[text]], StatementType.GENERAL_INFO)
                ])
                self.assertEqual(detection.regime, expected)
                self.assertEqual(detection.confidence, DetectionConfidence.HIGH)

    def test_multiple_explicit_references_are_ambiguous(self):
        detection = detect_accounting_regime([
            ExtractedTable(
                1,
                [["Thông tư 200/2014/TT-BTC"], ["Thông tư 99/2025/TT-BTC"]],
                StatementType.GENERAL_INFO,
            )
        ])

        self.assertEqual(detection.regime, AccountingRegime.AMBIGUOUS)

    def test_two_tt200_form_markers_produce_medium_confidence(self):
        detection = detect_accounting_regime([
            ExtractedTable(1, [["Mẫu số B01-DN"]], StatementType.BALANCE_SHEET_ASSETS),
            ExtractedTable(2, [["Mẫu số B02-DN"]], StatementType.INCOME_STATEMENT),
        ])

        self.assertEqual(detection.regime, AccountingRegime.TT200_2014)
        self.assertEqual(detection.confidence, DetectionConfidence.MEDIUM)

    def test_two_tt133_form_markers_produce_medium_confidence(self):
        detection = detect_accounting_regime([
            ExtractedTable(1, [["Mẫu số B01a-DNN"]], StatementType.BALANCE_SHEET_ASSETS),
            ExtractedTable(2, [["Mẫu số B02-DNN"]], StatementType.INCOME_STATEMENT),
        ])

        self.assertEqual(detection.regime, AccountingRegime.TT133_2016)
        self.assertEqual(detection.confidence, DetectionConfidence.MEDIUM)

    def test_unknown_document_is_not_assigned_tt200(self):
        detection = detect_accounting_regime([
            ExtractedTable(1, [["Báo cáo tài chính"]], StatementType.UNKNOWN)
        ])

        self.assertEqual(detection.regime, AccountingRegime.UNKNOWN)
        self.assertEqual(detection.confidence, DetectionConfidence.NONE)


if __name__ == "__main__":
    unittest.main()

