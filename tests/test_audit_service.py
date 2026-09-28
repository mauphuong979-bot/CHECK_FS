import unittest
from unittest.mock import patch

from src.reporting.excel_report import NOTE_SORT_SCORE, SUMMARY_SORT_STATUS
from src.domain.audit_rules import EMPTY_RULE_PACK, TT200_RULE_PACK
from src.domain.models import AccountingRegime, AccountingRegimeDetection, DetectionConfidence
from src.domain.models import ExtractedTable, StatementType, WordCommentCategory, WordCommentTarget
from src.services.audit_service import create_audit_outputs, create_audit_report


class AuditServiceSortOptionsTest(unittest.TestCase):
    @patch("src.services.audit_service.build_audit_workbook", return_value=b"workbook")
    @patch("src.services.audit_service.analyze_file", return_value=[])
    def test_passes_sort_options_to_reporter(self, analyze_file, build_audit_workbook):
        tables, workbook = create_audit_report(
            "bao_cao.docx",
            summary_sort=SUMMARY_SORT_STATUS,
            note_reconciliation_sort=NOTE_SORT_SCORE,
        )

        self.assertEqual(tables, [])
        self.assertEqual(workbook, b"workbook")
        analyze_file.assert_called_once_with("bao_cao.docx")
        build_audit_workbook.assert_called_once_with(
            [],
            rule_pack=EMPTY_RULE_PACK,
            summary_sort=SUMMARY_SORT_STATUS,
            note_reconciliation_sort=NOTE_SORT_SCORE,
        )

    @patch("src.services.audit_service.build_commented_docx", side_effect=RuntimeError("failed"))
    @patch(
        "src.services.audit_service.build_audit_workbook_with_comment_targets",
        return_value=(b"workbook", []),
    )
    @patch("src.services.audit_service.analyze_file", return_value=[])
    def test_word_failure_does_not_block_excel(self, analyze_file, build_workbook, build_word):
        outputs = create_audit_outputs("bao_cao.docx")

        self.assertEqual(outputs.workbook_bytes, b"workbook")
        self.assertIsNone(outputs.commented_docx_bytes)
        self.assertTrue(outputs.word_comment_failed)
        build_word.assert_called_once()

    @patch("src.services.audit_service.build_commented_docx", return_value=b"word")
    @patch(
        "src.services.audit_service.detect_accounting_regime",
        return_value=AccountingRegimeDetection(
            AccountingRegime.TT200_2014,
            DetectionConfidence.HIGH,
        ),
    )
    @patch(
        "src.services.audit_service.build_audit_workbook_with_comment_targets",
        return_value=(b"workbook", []),
    )
    @patch("src.services.audit_service.analyze_file", return_value=[])
    def test_applies_tt200_pack_only_after_detection(
        self, analyze_file, build_workbook, detector, build_word
    ):
        outputs = create_audit_outputs("bao_cao.docx")

        self.assertEqual(outputs.applied_rule_pack, TT200_RULE_PACK)
        self.assertEqual(build_workbook.call_args.kwargs["rule_pack"], TT200_RULE_PACK)

    @patch("src.services.audit_service.build_commented_docx", return_value=b"word")
    @patch(
        "src.services.audit_service.build_audit_workbook_with_comment_targets",
        return_value=(
            b"workbook",
            [
                WordCommentTarget(
                    1, 2, 1, "Kiểm tra số học: Có chênh lệch.",
                    WordCommentCategory.DIFFERENCE,
                ),
                WordCommentTarget(
                    2, 0, 0, "Kiểm tra số học: Cần xem xét thủ công.",
                    WordCommentCategory.REVIEW,
                ),
            ],
        ),
    )
    @patch(
        "src.services.audit_service.analyze_file",
        return_value=[
            ExtractedTable(1, [["Tổng cộng", "100"]], StatementType.NOTE, "Thuế"),
            ExtractedTable(2, [["Nội dung"]], StatementType.UNKNOWN, "Bảng khác"),
        ],
    )
    def test_builds_prioritized_attention_items(self, analyze_file, build_workbook, build_word):
        outputs = create_audit_outputs("bao_cao.docx")

        self.assertEqual([item.severity for item in outputs.attention_items], ["Sai lệch", "Cần xem xét"])
        self.assertEqual(outputs.attention_items[0].table_title, "Thuế")
        self.assertEqual(outputs.attention_items[0].table_type, "Thuyết minh")
        self.assertEqual(outputs.attention_items[1].table_title, "Chưa phân loại")
        self.assertEqual(outputs.attention_items[1].table_type, "Chưa phân loại")


if __name__ == "__main__":
    unittest.main()
