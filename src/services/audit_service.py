from __future__ import annotations

from dataclasses import dataclass, replace
import logging
from pathlib import Path

from src.domain.accounting_regime_detector import detect_accounting_regime
from src.domain.audit_rules import EMPTY_RULE_PACK, RulePack, get_rule_pack
from src.io.docx_reader import read_docx_tables, read_docx_text_fragments
from src.domain.models import (
    AccountingRegimeDetection,
    FindingDetail,
    StatementType,
    TableCheckResult,
    WordCommentCategory,
    WordCommentTarget,
)
from src.extraction.table_currency import resolve_table_currencies
from src.io.file_detector import detect_extension
from src.normalization.number_parser import parse_accounting_number
from src.reporting.excel_report import (
    NOTE_SORT_SCORE,
    NOTE_SORT_STATUS,
    NOTE_SORT_TABLE_INDEX,
    SUMMARY_SORT_ISSUES_DESC,
    SUMMARY_SORT_STATUS,
    SUMMARY_SORT_TABLE_INDEX,
    build_audit_report,
    build_audit_workbook,
    table_content_label,
    table_type_label,
)
from src.reporting.word_comment_report import build_commented_docx


LOGGER = logging.getLogger(__name__)


class UnsupportedFileError(ValueError):
    pass


CHECKED_TABLE_STATUSES = {"Có kiểm tra", "Có rule nghiệp vụ"}
NOT_APPLICABLE_TABLE_STATUSES = {"Không kiểm tra"}
SEVERITY_DIFFERENCE = "Sai lệch"
SEVERITY_REVIEW = "Cần xem xét"


@dataclass(frozen=True)
class AuditCoverage:
    """Phạm vi kiểm tra tự động, để UI không kết luận "khớp" cho bảng chưa được kiểm tra."""

    total_tables: int = 0
    checked_tables: int = 0
    review_tables: int = 0
    not_applicable_tables: int = 0

    @property
    def unchecked_financial_tables(self) -> int:
        return self.review_tables


@dataclass(frozen=True)
class AuditOutputs:
    tables: list
    workbook_bytes: bytes
    commented_docx_bytes: bytes | None = None
    word_comment_failed: bool = False
    attention_items: tuple["AuditAttentionItem", ...] = ()
    regime_detection: AccountingRegimeDetection | None = None
    applied_rule_pack: RulePack = EMPTY_RULE_PACK
    coverage: AuditCoverage = AuditCoverage()


@dataclass(frozen=True)
class AuditAttentionItem:
    severity: str
    table_index: int
    table_title: str
    table_type: str
    message: str
    location: str = ""
    detail: FindingDetail | None = None
    statement_label: str = ""
    currency: str = ""


def analyze_file(path: str | Path):
    extension = detect_extension(str(path))
    if extension != ".docx":
        raise UnsupportedFileError("Hiện MVP chỉ xử lý tệp DOCX. Các định dạng khác sẽ được bổ sung sau.")
    return read_docx_tables(path)


def create_audit_report(
    path: str | Path,
    *,
    summary_sort: str = SUMMARY_SORT_TABLE_INDEX,
    note_reconciliation_sort: str = NOTE_SORT_TABLE_INDEX,
) -> tuple[list, bytes]:
    tables = analyze_file(path)
    detection = _detect_regime(path, tables)
    rule_pack = get_rule_pack(detection.regime) or EMPTY_RULE_PACK
    workbook_bytes = build_audit_workbook(
        tables,
        rule_pack=rule_pack,
        summary_sort=summary_sort,
        note_reconciliation_sort=note_reconciliation_sort,
    )
    return tables, workbook_bytes


def create_audit_outputs(
    path: str | Path,
    *,
    summary_sort: str = SUMMARY_SORT_TABLE_INDEX,
    note_reconciliation_sort: str = NOTE_SORT_TABLE_INDEX,
) -> AuditOutputs:
    tables = analyze_file(path)
    detection = _detect_regime(path, tables)
    rule_pack = get_rule_pack(detection.regime) or EMPTY_RULE_PACK
    report = build_audit_report(
        tables,
        rule_pack=rule_pack,
        summary_sort=summary_sort,
        note_reconciliation_sort=note_reconciliation_sort,
    )
    attention_items = _build_attention_items(tables, report.comment_targets)
    coverage = _build_coverage(report.table_results)
    try:
        commented_docx_bytes = build_commented_docx(path, report.comment_targets)
    except Exception as exc:
        # Lỗi tạo Word tại ranh giới này không được làm mất báo cáo Excel.
        # Chỉ ghi loại lỗi, không ghi nội dung để tránh lộ dữ liệu báo cáo.
        LOGGER.warning("Không tạo được file Word có comment: %s", type(exc).__name__)
        commented_docx_bytes = None
    return AuditOutputs(
        tables=tables,
        workbook_bytes=report.workbook_bytes,
        commented_docx_bytes=commented_docx_bytes,
        word_comment_failed=commented_docx_bytes is None,
        attention_items=attention_items,
        regime_detection=detection,
        applied_rule_pack=rule_pack,
        coverage=coverage,
    )


def _build_coverage(results: tuple[TableCheckResult, ...]) -> AuditCoverage:
    checked = sum(1 for result in results if result.status in CHECKED_TABLE_STATUSES)
    not_applicable = sum(
        1 for result in results if result.status in NOT_APPLICABLE_TABLE_STATUSES
    )
    return AuditCoverage(
        total_tables=len(results),
        checked_tables=checked,
        review_tables=len(results) - checked - not_applicable,
        not_applicable_tables=not_applicable,
    )


def _detect_regime(path: str | Path, tables: list) -> AccountingRegimeDetection:
    source_path = Path(path)
    text_fragments = read_docx_text_fragments(source_path) if source_path.is_file() else ()
    return detect_accounting_regime(tables, text_fragments)


STATEMENT_SHORT_LABELS = {
    StatementType.BALANCE_SHEET_ASSETS: "BS",
    StatementType.BALANCE_SHEET_EQUITY: "BS",
    StatementType.INCOME_STATEMENT: "PL",
    StatementType.CASH_FLOW: "CF",
    StatementType.NOTE: "TM",
}


def _build_attention_items(
    tables: list,
    comment_targets: list[WordCommentTarget],
) -> tuple[AuditAttentionItem, ...]:
    table_by_index = {table.index: table for table in tables}
    currency_by_table = resolve_table_currencies(tables)
    items: list[AuditAttentionItem] = []
    seen: set[tuple[int, str, str]] = set()
    for target in comment_targets:
        table = table_by_index.get(target.table_index)
        if table is None:
            continue
        severity = (
            SEVERITY_DIFFERENCE
            if target.category == WordCommentCategory.DIFFERENCE
            else SEVERITY_REVIEW
        )
        identity = (target.table_index, severity, target.message)
        if identity in seen:
            continue
        seen.add(identity)
        items.append(
            AuditAttentionItem(
                severity=severity,
                table_index=table.index,
                table_title=table_content_label(table),
                table_type=table_type_label(table),
                message=target.message,
                location=_target_location(table, target),
                detail=target.detail,
                statement_label=STATEMENT_SHORT_LABELS.get(table.statement_type, "Khác"),
                currency=currency_by_table.get(table.index) or "",
            )
        )
    merged = _merge_periods_without_amounts(items)
    return tuple(
        sorted(
            merged,
            key=lambda item: (
                0 if item.severity == SEVERITY_DIFFERENCE else 1,
                item.table_index,
                item.message,
            ),
        )
    )


def _merge_periods_without_amounts(
    items: list[AuditAttentionItem],
) -> list[AuditAttentionItem]:
    """Gộp cùng một vấn đề không có số tiền xuất hiện ở cả kỳ này và kỳ trước thành một dòng.

    Phát hiện có số chênh lệch luôn giữ riêng từng kỳ vì số liệu khác nhau.
    """
    merged: list[AuditAttentionItem] = []
    index_by_key: dict[tuple, int] = {}
    for item in items:
        detail = item.detail
        if detail is None or detail.difference is not None or not detail.period:
            merged.append(item)
            continue
        key = (item.severity, item.table_index, item.location, detail.check, detail.note)
        if key not in index_by_key:
            index_by_key[key] = len(merged)
            merged.append(item)
            continue
        position = index_by_key[key]
        existing = merged[position]
        periods = list(dict.fromkeys([*existing.detail.period.split(", "), detail.period]))
        merged[position] = replace(
            existing,
            detail=replace(existing.detail, period=", ".join(periods)),
        )
    return merged


def _target_location(table, target: WordCommentTarget) -> str:
    """Nhãn dòng của ô được comment để kiểm toán viên định vị nhanh trên Word."""
    if not 0 <= target.row_index < len(table.rows):
        return ""
    row = table.rows[target.row_index]
    label = next(
        (value.strip() for value in row[:2] if value and value.strip() and not _is_number(value)),
        "",
    )
    # Bỏ ký hiệu gạch đầu dòng của chỉ tiêu con (ví dụ "-Lợi nhuận…") để nhãn gọn hơn.
    return label.lstrip("-–—•+ ").strip()[:80]


def _is_number(value: str) -> bool:
    return parse_accounting_number(value) is not None
