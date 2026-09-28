from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from src.domain.accounting_regime_detector import detect_accounting_regime
from src.domain.audit_rules import EMPTY_RULE_PACK, RulePack, get_rule_pack
from src.io.docx_reader import read_docx_tables, read_docx_text_fragments
from src.domain.models import AccountingRegimeDetection, WordCommentCategory, WordCommentTarget
from src.io.file_detector import detect_extension
from src.reporting.excel_report import (
    NOTE_SORT_SCORE,
    NOTE_SORT_STATUS,
    NOTE_SORT_TABLE_INDEX,
    SUMMARY_SORT_ISSUES_DESC,
    SUMMARY_SORT_STATUS,
    SUMMARY_SORT_TABLE_INDEX,
    build_audit_workbook,
    build_audit_workbook_with_comment_targets,
    table_content_label,
    table_type_label,
)
from src.reporting.word_comment_report import build_commented_docx


class UnsupportedFileError(ValueError):
    pass


@dataclass(frozen=True)
class AuditOutputs:
    tables: list
    workbook_bytes: bytes
    commented_docx_bytes: bytes | None = None
    word_comment_failed: bool = False
    attention_items: tuple["AuditAttentionItem", ...] = ()
    regime_detection: AccountingRegimeDetection | None = None
    applied_rule_pack: RulePack = EMPTY_RULE_PACK


@dataclass(frozen=True)
class AuditAttentionItem:
    severity: str
    table_index: int
    table_title: str
    table_type: str
    message: str


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
    workbook_bytes, comment_targets = build_audit_workbook_with_comment_targets(
        tables,
        rule_pack=rule_pack,
        summary_sort=summary_sort,
        note_reconciliation_sort=note_reconciliation_sort,
    )
    attention_items = _build_attention_items(tables, comment_targets)
    try:
        commented_docx_bytes = build_commented_docx(path, comment_targets)
    except Exception:
        # Lỗi tạo Word tại ranh giới này không được làm mất báo cáo Excel.
        return AuditOutputs(
            tables=tables,
            workbook_bytes=workbook_bytes,
            word_comment_failed=True,
            attention_items=attention_items,
            regime_detection=detection,
            applied_rule_pack=rule_pack,
        )
    return AuditOutputs(
        tables=tables,
        workbook_bytes=workbook_bytes,
        commented_docx_bytes=commented_docx_bytes,
        attention_items=attention_items,
        regime_detection=detection,
        applied_rule_pack=rule_pack,
    )


def _detect_regime(path: str | Path, tables: list) -> AccountingRegimeDetection:
    source_path = Path(path)
    text_fragments = read_docx_text_fragments(source_path) if source_path.is_file() else ()
    return detect_accounting_regime(tables, text_fragments)


def _build_attention_items(
    tables: list,
    comment_targets: list[WordCommentTarget],
) -> tuple[AuditAttentionItem, ...]:
    table_by_index = {table.index: table for table in tables}
    items: list[AuditAttentionItem] = []
    seen: set[tuple[int, str, str]] = set()
    for target in comment_targets:
        table = table_by_index.get(target.table_index)
        if table is None:
            continue
        severity = (
            "Sai lệch"
            if target.category == WordCommentCategory.DIFFERENCE
            else "Cần xem xét"
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
            )
        )
    return tuple(
        sorted(
            items,
            key=lambda item: (
                0 if item.severity == "Sai lệch" else 1,
                item.table_index,
                item.message,
            ),
        )
    )
