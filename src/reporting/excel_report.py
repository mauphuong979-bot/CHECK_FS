from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO
from pathlib import Path
import re

from openpyxl import Workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from src.domain.audit_rules import (
    ArithmeticRule,
    CrossPeriodRule,
    RulePack,
    SignRule,
    TT200_RULE_PACK,
)
from src.domain.arithmetic_checker import calculate_same_table_arithmetic
from src.domain.models import (
    ExtractedTable,
    NoteMatchResult,
    StatementType,
    TableCheckResult,
    TableValueStatus,
    WordCommentCategory,
    WordCommentTarget,
)
from src.extraction.table_layout import (
    TableLayout,
    detect_code_column as _detect_code_col,
    detect_period_columns_from_headers,
    detect_table_layout as _table_layout,
    is_note_column as _is_note_column,
    numeric_columns_for as _numeric_columns,
    select_period_column as _period_col,
)
from src.extraction.table_values import (
    DEFAULT_EXCEL_DATA_START_ROW,
    code_excel_row_lookup as _code_row_lookup,
    normalize_statement_code as _normalize_code,
    table_value_lookup as _table_value_lookup,
    table_value_state_lookup,
)
from src.normalization.number_parser import parse_accounting_number
from src.normalization.text_cleaner import clean_text, normalized_key
from src.reconciliation.note_matcher import (
    is_parent_detail_with_maturity_split,
    is_tax_rollforward_note,
    reconcile_note_tables,
)
from src.reporting.excel_styles import (
    CHECK_VALUE_FILL,
    COLOR_BLUE,
    COLOR_HYPERLINK,
    COLOR_MUTED_TEXT,
    COLOR_NAVY,
    COLOR_TEXT,
    COLOR_WHITE,
    COLOR_YELLOW,
    GROUP_HEADER_FILL,
    HEADER_FILL,
    LINKED_VALUE_FILL,
    STATUS_FILLS,
    TABLE_STATUS_FILLS,
    THIN_BORDER,
    TITLE_FILL,
    WARNING_FILL,
)


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


def _excel_check_formula(expression: str) -> str:
    return f"=ROUND({expression},{CHECK_ROUND_DIGITS})"


def _rounded_check_value(value: Decimal | float | int) -> Decimal:
    return Decimal(str(value)).quantize(CHECK_ROUND_QUANTUM, rounding=ROUND_HALF_UP)


def _has_check_difference(value: Decimal | float | int) -> bool:
    return _rounded_check_value(value) != Decimal(0)


def _format_vietnamese_number(value: Decimal | float | int) -> str:
    rounded = _rounded_check_value(abs(value))
    formatted = f"{rounded:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return formatted[:-3] if formatted.endswith(",00") else formatted


def _arithmetic_difference_message(value: Decimal | float | int) -> str:
    return f"Kiểm tra số học: chênh lệch {_format_vietnamese_number(value)}."


def _mark_arithmetic_difference(cell, value: Decimal | float | int) -> None:
    cell = _merged_anchor_cell(cell)
    cell.fill = WARNING_FILL
    cell.comment = Comment(_arithmetic_difference_message(value), "CHECK_FS_RULE")


def _merged_anchor_cell(cell):
    """Trả về ô neo có thể ghi của một vùng merge, nếu ``cell`` là ô phụ."""
    if not isinstance(cell, MergedCell):
        return cell
    for merged_range in cell.parent.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return cell.parent.cell(
                row=merged_range.min_row,
                column=merged_range.min_col,
            )
    return cell


def build_audit_workbook(
    tables: list[ExtractedTable],
    *,
    rule_pack: RulePack = TT200_RULE_PACK,
    summary_sort: str = SUMMARY_SORT_TABLE_INDEX,
    note_reconciliation_sort: str = NOTE_SORT_TABLE_INDEX,
) -> bytes:
    workbook_bytes, _ = build_audit_workbook_with_comment_targets(
        tables,
        rule_pack=rule_pack,
        summary_sort=summary_sort,
        note_reconciliation_sort=note_reconciliation_sort,
    )
    return workbook_bytes


def build_audit_workbook_with_comment_targets(
    tables: list[ExtractedTable],
    *,
    rule_pack: RulePack = TT200_RULE_PACK,
    summary_sort: str = SUMMARY_SORT_TABLE_INDEX,
    note_reconciliation_sort: str = NOTE_SORT_TABLE_INDEX,
) -> tuple[bytes, list[WordCommentTarget]]:
    _validate_sort_option(summary_sort, VALID_SUMMARY_SORTS, "00_Tong_hop")
    workbook = Workbook()
    summary = workbook.active
    summary.title = "00_Tong_hop"

    sheet_names = {table.index: _sheet_name(table) for table in tables}
    raw_note_matches = (
        reconcile_note_tables(tables)
        if rule_pack.regime == TT200_RULE_PACK.regime
        else []
    )
    note_matches = _ordered_note_matches(raw_note_matches)
    note_matches_by_table: dict[int, list[NoteMatchResult]] = {}
    for match in note_matches:
        note_matches_by_table.setdefault(match.note_table_index, []).append(match)
    maturity_parent_indexes = {
        table.index
        for table in tables
        if is_parent_detail_with_maturity_split(table, tables)
    }
    results: list[TableCheckResult] = []
    worksheets_by_index = {}
    for table_position, table in enumerate(tables):
        peer_tables = _nearest_statement_tables(tables, table)
        peer_sheets = {peer.statement_type: sheet_names[peer.index] for peer in peer_tables.values()}
        peer_layouts = {peer.statement_type: _table_layout(peer) for peer in peer_tables.values()}
        peer_code_rows = {
            peer.statement_type: _code_row_lookup(
                peer,
                _table_layout(peer),
                data_start_row=DATA_START_ROW,
            )
            for peer in peer_tables.values()
        }
        peer_values = {
            peer.statement_type: _table_value_lookup(peer, _table_layout(peer))
            for peer in peer_tables.values()
        }
        sheet_name = sheet_names[table.index]
        worksheet = workbook.create_sheet(sheet_name)
        worksheets_by_index[table.index] = worksheet
        result = _write_table_sheet(
            worksheet,
            table,
            peer_sheets,
            peer_layouts,
            peer_code_rows,
            peer_values,
            rule_pack=rule_pack,
            previous_table=tables[table_position - 1] if table_position else None,
            previous_sheet_name=(
                sheet_names[tables[table_position - 1].index]
                if table_position
                else ""
            ),
        )
        if table.statement_type == StatementType.NOTE:
            _write_note_reconciliation_detail(
                worksheet,
                note_matches_by_table.get(table.index, []),
                sheet_names,
                table,
                maturity_split_parent=table.index in maturity_parent_indexes,
            )
        results.append(result)

    summary_tables = _sort_summary_tables(tables, results, summary_sort)
    result_by_index = {result.table_index: result for result in results}
    detail_tables = sorted(
        tables,
        key=lambda table: _detail_sheet_sort_key(table, result_by_index),
    )
    detail_sheets = [worksheets_by_index[table.index] for table in detail_tables]
    workbook._sheets = [summary, *detail_sheets]
    _write_summary(
        summary,
        summary_tables,
        results,
        note_matches_by_table,
        rule_pack,
        maturity_parent_indexes=maturity_parent_indexes,
    )
    _update_summary_backlinks(worksheets_by_index, summary_tables)
    _write_detail_navigation(worksheets_by_index, detail_tables)
    comment_targets = _build_word_comment_targets(
        tables,
        results,
        worksheets_by_index,
        raw_note_matches,
    )
    _apply_sheet_tab_colors(
        summary,
        worksheets_by_index,
        results,
        raw_note_matches,
        comment_targets=comment_targets,
    )

    stream = BytesIO()
    workbook.save(stream)
    stream.seek(0)
    return stream.getvalue(), comment_targets


def _build_word_comment_targets(
    tables: list[ExtractedTable],
    results: list[TableCheckResult],
    worksheets_by_index: dict[int, object],
    note_matches: list[NoteMatchResult],
) -> list[WordCommentTarget]:
    targets: list[WordCommentTarget] = []
    table_by_index = {table.index: table for table in tables}

    for table in tables:
        worksheet = worksheets_by_index[table.index]
        for row_index in range(table.row_count):
            for column_index in range(table.column_count):
                cell = worksheet.cell(
                    row=DATA_START_ROW + row_index,
                    column=column_index + 1,
                )
                if cell.fill.fgColor.rgb not in {COLOR_YELLOW, f"00{COLOR_YELLOW}"}:
                    continue
                # Comment có sẵn tại vùng dữ liệu hiện chỉ dùng cho cảnh báo định dạng số.
                # Các cảnh báo này vẫn xuất hiện trong Excel nhưng không chuyển sang Word.
                if cell.comment is not None:
                    if cell.comment.author != "CHECK_FS_RULE":
                        continue
                    message = cell.comment.text
                else:
                    message = "Kiểm tra số học: Có chênh lệch."

                msg_lower = message.lower()
                is_review_only = (
                    "cần xác minh" in msg_lower
                    or "thiếu hoặc không hợp lệ" in msg_lower
                    or "thiếu" in msg_lower
                    or "không hợp lệ" in msg_lower
                )
                category = (
                    WordCommentCategory.REVIEW
                    if is_review_only
                    else WordCommentCategory.DIFFERENCE
                )

                targets.append(
                    WordCommentTarget(
                        table.index,
                        row_index,
                        column_index,
                        message,
                        category,
                    )
                )
    for match in note_matches:
        rich_match = match.match_type in {
            "Tên dòng và Thuyết minh",
            "Tiêu đề kỳ hạn và Thuyết minh",
            "Mã Thuyết minh + tên + Tổng cộng",
            "Tên dòng + PL",
            "Tên dòng + mã BS",
            "Tên dòng + mã CF",
            "Vốn đã góp + mã BS",
            "Tên bảng + Tổng cộng",
        }
        if match.status == "Matched":
            continue
        if (
            match.status == "Statement not found"
            and not match.statement_code
            and not rich_match
        ):
            continue
        table = table_by_index.get(match.note_table_index)
        if table is None:
            continue
        references: list[str] = []
        if match.status == "Difference":
            references = [
                reference
                for reference, difference in (
                    (match.note_current_cell, match.current_difference),
                    (match.note_prior_cell, match.prior_difference),
                )
                if reference
                and difference is not None
                and _has_check_difference(difference)
            ]
            if not references:
                references = [match.note_current_cell or match.note_prior_cell]
        elif rich_match:
            references = [match.note_current_cell, match.note_prior_cell]
            if (
                match.status == "Statement not found"
                and match.match_type == "Tên dòng + mã BS"
                and match.note_current == Decimal(0)
                and match.note_prior == Decimal(0)
            ):
                references = [match.note_current_cell or match.note_prior_cell]
        else:
            references = [match.note_current_cell or match.note_prior_cell]
        row_index, column_index = _first_meaningful_source_cell(table)
        for reference in {reference for reference in references if reference} or {""}:
            if reference:
                source_location = _source_location_from_excel_reference(reference, table)
                if source_location is not None:
                    row_index, column_index = source_location
            targets.append(
                WordCommentTarget(
                    table.index,
                    row_index,
                    column_index,
                    _word_note_comment_message(match, reference),
                    (
                        WordCommentCategory.DIFFERENCE
                        if match.status == "Difference"
                        else WordCommentCategory.REVIEW
                    ),
                )
            )
    return targets


def _source_location_from_excel_reference(
    reference: str,
    table: ExtractedTable,
) -> tuple[int, int] | None:
    match = re.fullmatch(r"([A-Z]+)(\d+)", reference.upper())
    if not match:
        return None
    column_index = 0
    for character in match.group(1):
        column_index = column_index * 26 + ord(character) - 64
    row_index = int(match.group(2)) - DATA_START_ROW
    column_index -= 1
    if 0 <= row_index < table.row_count and 0 <= column_index < table.column_count:
        return _merged_anchor_location(table, row_index, column_index)
    return None


def _merged_anchor_location(
    table: ExtractedTable,
    row_index: int,
    column_index: int,
) -> tuple[int, int]:
    for start_row, start_col, end_row, end_col in table.merged_ranges:
        if start_row <= row_index <= end_row and start_col <= column_index <= end_col:
            return start_row, start_col
    return row_index, column_index


def _merged_anchor_reference(table: ExtractedTable, reference: str) -> str:
    source_location = _source_location_from_excel_reference(reference, table)
    if source_location is None:
        return reference
    row_index, column_index = source_location
    return f"{get_column_letter(column_index + 1)}{row_index + DATA_START_ROW}"


def _first_meaningful_source_cell(table: ExtractedTable) -> tuple[int, int]:
    for row_index, row in enumerate(table.rows):
        for column_index, value in enumerate(row):
            if clean_text(value):
                return row_index, column_index
    return 0, 0


def _word_note_comment_message(match: NoteMatchResult, reference: str = "") -> str:
    status = NOTE_STATUS_LABELS.get(match.status, match.status)
    referenced_differences = [
        difference
        for cell_reference, difference in (
            (match.note_current_cell, match.current_difference),
            (match.note_prior_cell, match.prior_difference),
        )
        if reference and reference == cell_reference and difference is not None
    ]
    difference = next(
        (value for value in referenced_differences if _has_check_difference(value)),
        referenced_differences[0] if referenced_differences else match.current_difference,
    )
    if match.match_type not in {
        "Tên dòng và Thuyết minh",
        "Tiêu đề kỳ hạn và Thuyết minh",
        "Mã Thuyết minh + tên + Tổng cộng",
        "Tên dòng + PL",
        "Tên dòng + mã BS",
        "Tên dòng + mã CF",
        "Vốn đã góp + mã BS",
        "Tên bảng + Tổng cộng",
    }:
        if difference is not None and _has_check_difference(difference):
            formatted = _format_vietnamese_number(difference)
            return f"Đối chiếu TM–BCTC: Có chênh lệch {formatted} VND."
        return f"Đối chiếu TM–BCTC: {status}."
    source = match.source or "BCTC"
    target = match.item_name or match.note_title or "chỉ tiêu phù hợp"
    if match.status == "Statement not found":
        return (
            f"Đối chiếu TM–{source}: Chưa tìm thấy chỉ tiêu phù hợp cho “{target}”. "
            "Cần kiểm toán viên xác minh."
        )
    code = f", mã số {match.statement_code}" if match.statement_code else ""
    prefix = f"Đối chiếu TM–{source}: Tìm thấy tại “{target}”{code}."
    if difference is not None and _has_check_difference(difference):
        formatted = _format_vietnamese_number(difference)
        return f"{prefix} Có chênh lệch {formatted} VND."
    if match.status == "Note value not found":
        return f"{prefix} Chưa tìm thấy số liệu TM tương ứng."
    return f"{prefix} Số liệu khớp."


def save_audit_workbook(
    tables: list[ExtractedTable],
    output_path: str | Path,
    *,
    summary_sort: str = SUMMARY_SORT_TABLE_INDEX,
    note_reconciliation_sort: str = NOTE_SORT_TABLE_INDEX,
) -> None:
    Path(output_path).write_bytes(
        build_audit_workbook(
            tables,
            summary_sort=summary_sort,
            note_reconciliation_sort=note_reconciliation_sort,
        )
    )


def _sheet_name(table: ExtractedTable) -> str:
    suffix_by_type = {
        StatementType.BALANCE_SHEET_ASSETS: "BS_TS",
        StatementType.BALANCE_SHEET_EQUITY: "BS_NV",
        StatementType.INCOME_STATEMENT: "PL",
        StatementType.CASH_FLOW: "CF",
        StatementType.NOTE: "TM",
        StatementType.GENERAL_INFO: "TT",
        StatementType.SIGNATURE: "CK",
        StatementType.UNKNOWN: "KHAC",
    }
    suffix = suffix_by_type.get(table.statement_type, "KHAC")
    return f"T{table.index:03d}_{suffix}"[:31]


def table_content_label(table: ExtractedTable) -> str:
    """Nhãn nội dung bảng dùng thống nhất cho Excel và giao diện."""
    if _is_related_party_non_total_table(table):
        return "Bên liên quan"
    if table.statement_type == StatementType.NOTE:
        raw_title = clean_text(table.title_hint or table.statement_type.value)
        title_key = normalized_key(raw_title)
        if title_key in {"ngan han", "dai han", "khac"} and table.context_hints:
            for ctx in table.context_hints:
                ctx_cleaned = clean_text(ctx)
                ctx_key = normalized_key(ctx_cleaned)
                if ctx_cleaned and ctx_key not in {"ngan han", "dai han", "khac"}:
                    return f"{ctx_cleaned} - {raw_title}"
        return raw_title
    return table.statement_type.value


def table_type_label(table: ExtractedTable) -> str:
    """Nhãn loại bảng dùng thống nhất cho Excel và giao diện."""
    if _is_related_party_relationship_table(table):
        return "Giao dịch với các bên liên quan"
    return table.statement_type.value


def _sort_summary_tables(
    tables: list[ExtractedTable],
    results: list[TableCheckResult],
    sort_option: str,
) -> list[ExtractedTable]:
    result_by_index = {result.table_index: result for result in results}
    if sort_option == SUMMARY_SORT_ISSUES_DESC:
        return sorted(
            tables,
            key=lambda table: (-result_by_index[table.index].issue_count, table.index),
        )
    if sort_option == SUMMARY_SORT_STATUS:
        return sorted(
            tables,
            key=lambda table: (normalized_key(result_by_index[table.index].status), table.index),
        )
    return sorted(tables, key=lambda table: table.index)


def _detail_sheet_sort_key(
    table: ExtractedTable,
    result_by_index: dict[int, TableCheckResult],
) -> tuple[int, int]:
    result = result_by_index.get(table.index)
    has_checks_or_issues = (
        result is not None
        and (result.check_count > 0 or result.issue_count > 0 or result.status == "Có kiểm tra")
    )

    primary_order = {
        StatementType.BALANCE_SHEET_ASSETS: 10,
        StatementType.BALANCE_SHEET_EQUITY: 11,
        StatementType.INCOME_STATEMENT: 12,
        StatementType.CASH_FLOW: 13,
    }

    if table.statement_type in primary_order:
        bucket = primary_order[table.statement_type]
    elif table.statement_type == StatementType.NOTE:
        if has_checks_or_issues:
            bucket = 20
        elif _is_non_arithmetic_information_table(table):
            bucket = 35
        else:
            bucket = 30
    elif table.statement_type == StatementType.GENERAL_INFO:
        bucket = 40
    elif table.statement_type == StatementType.SIGNATURE:
        bucket = 50
    else:
        if has_checks_or_issues:
            bucket = 20
        else:
            bucket = 60

    return (bucket, table.index)


def _apply_sheet_tab_colors(
    summary_sheet,
    worksheets_by_index: dict[int, object],
    results: list[TableCheckResult],
    raw_note_matches: list[NoteMatchResult],
    comment_targets: list[WordCommentTarget] | None = None,
) -> None:
    summary_sheet.sheet_properties.tabColor = COLOR_NAVY
    result_by_index = {res.table_index: res for res in results}

    note_matches_by_table: dict[int, list[NoteMatchResult]] = {}
    for match in raw_note_matches:
        note_matches_by_table.setdefault(match.note_table_index, []).append(match)
        if match.statement_table_index is not None:
            note_matches_by_table.setdefault(match.statement_table_index, []).append(match)

    comment_targets_by_table: dict[int, list[WordCommentTarget]] = {}
    if comment_targets:
        for target in comment_targets:
            comment_targets_by_table.setdefault(target.table_index, []).append(target)

    TAB_COLOR_RED = "C00000"
    TAB_COLOR_YELLOW = "FFC000"
    TAB_COLOR_GREEN = "70AD47"

    for table_index, worksheet in worksheets_by_index.items():
        result = result_by_index.get(table_index)
        matches = note_matches_by_table.get(table_index, [])
        targets = comment_targets_by_table.get(table_index, [])
        if result is None:
            continue

        has_target_difference = any(t.category == WordCommentCategory.DIFFERENCE for t in targets)
        has_target_review = any(t.category == WordCommentCategory.REVIEW for t in targets)

        has_difference = (
            (result.issue_count > 0 and result.status not in {"Cần xem xét", "Không kiểm tra"})
            or has_target_difference
            or result.status in {"Có chênh lệch", "Có sai lệch"}
            or any(m.status == "Difference" for m in matches)
        )
        needs_review = (
            not has_difference and (
                result.status == "Cần xem xét"
                or has_target_review
                or any(
                    m.status in {"Statement not found", "Note value not found"}
                    and bool(m.statement_code)
                    for m in matches
                )
            )
        )
        is_matched = (
            not has_difference
            and not needs_review
            and (result.check_count > 0 or result.status in {"Khớp", "Có kiểm tra", "Có rule nghiệp vụ"})
        )

        if has_difference:
            worksheet.sheet_properties.tabColor = TAB_COLOR_RED
        elif needs_review:
            worksheet.sheet_properties.tabColor = TAB_COLOR_YELLOW
        elif is_matched:
            worksheet.sheet_properties.tabColor = TAB_COLOR_GREEN


def _validate_sort_option(value: str, valid_values: set[str], sheet_name: str) -> None:
    if value not in valid_values:
        raise ValueError(f"Tùy chọn sắp xếp không hợp lệ cho sheet {sheet_name}.")


def _write_table_sheet(
    worksheet,
    table: ExtractedTable,
    primary_sheets: dict[StatementType, str],
    primary_layouts: dict[StatementType, TableLayout],
    primary_code_rows: dict[StatementType, dict[str, int]],
    primary_values: dict[StatementType, dict[str, dict[int, float]]],
    *,
    rule_pack: RulePack,
    previous_table: ExtractedTable | None = None,
    previous_sheet_name: str = "",
) -> TableCheckResult:
    worksheet.freeze_panes = f"B{DATA_START_ROW}" if table.statement_type == StatementType.NOTE else f"A{DATA_START_ROW}"
    worksheet.sheet_view.showGridLines = True

    max_cols = table.column_count
    layout = _table_layout(table)
    numeric_columns = _numeric_columns(table)
    if table.statement_type == StatementType.NOTE:
        header_period_columns = _note_header_period_columns(table)
        numeric_columns.update(
            column for column in header_period_columns if column is not None
        )
    expected_number_style = _expected_number_style(table)
    format_issue_cells: list[str] = []

    title_cell = worksheet.cell(row=TITLE_ROW, column=2, value=table_content_label(table))
    title_cell.fill = TITLE_FILL
    title_cell.font = Font(bold=True, color=COLOR_WHITE, size=12)
    title_cell.border = THIN_BORDER
    title_cell.alignment = Alignment(horizontal="left", vertical="center")

    back_cell = worksheet.cell(row=INFO_ROW, column=2, value="Trở về Tổng hợp")
    back_cell.hyperlink = "#00_Tong_hop!A1"
    back_cell.font = Font(bold=True, color=COLOR_HYPERLINK, underline="single")
    back_cell.alignment = Alignment(horizontal="left", vertical="center")
    for col_idx in range(1, max_cols + 1):
        cell = worksheet.cell(row=HEADER_ROW, column=col_idx, value=f"Cột {col_idx}")
        cell.font = Font(bold=True, color=COLOR_WHITE)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER

    for row_idx, row in enumerate(table.rows, start=DATA_START_ROW):
        for col_idx in range(1, max_cols + 1):
            raw_value = row[col_idx - 1] if col_idx <= len(row) else ""
            parsed_number = parse_accounting_number(raw_value)
            cell = worksheet.cell(row=row_idx, column=col_idx)
            if parsed_number is not None:
                cell.value = float(parsed_number)
                cell.number_format = ACCOUNTING_NUMBER_FORMAT
                if _has_number_format_issue(raw_value, expected_number_style):
                    cell.fill = WARNING_FILL
                    cell.comment = Comment(
                        f"Định dạng số khác đa số trong bảng. Giá trị gốc: {clean_text(raw_value)}",
                        "CHECK_FS",
                    )
                    format_issue_cells.append(cell.coordinate)
            elif _is_dash_zero(raw_value) and col_idx in numeric_columns:
                cell.value = 0
                cell.number_format = ACCOUNTING_NUMBER_FORMAT
            else:
                cell.value = raw_value
            cell.border = THIN_BORDER

    if table.statement_type == StatementType.NOTE:
        _apply_note_source_presentation(worksheet, table)
    else:
        _highlight_total_rows(worksheet, table)

    check_start_col = max_cols + 1

    if table.statement_type in {StatementType.NOTE, StatementType.UNKNOWN} and (
        _is_non_arithmetic_information_table(table)
    ):
        _write_sheet_legend(worksheet)
        _apply_output_widths(worksheet, table, max_cols)
        result = TableCheckResult(
            table.index,
            "Không kiểm tra",
            0,
            0,
            "Bảng cung cấp thông tin mô tả, không có cấu trúc tổng để kiểm tra số học.",
        )
        return _finalize_result(worksheet, max_cols, result)

    business_result = _write_business_checks(
        worksheet,
        table,
        layout,
        primary_sheets,
        primary_layouts,
        primary_code_rows,
        primary_values,
        check_start_col,
        rule_pack,
    )
    if business_result.check_count:
        _style_check_area(worksheet, table, business_result.formula_cells, max_cols)
        _write_sheet_legend(worksheet)
        _apply_output_widths(worksheet, table, max_cols)
        return _finalize_result(worksheet, max_cols, business_result)

    formula_cells = _write_note_table_checks(
        worksheet,
        table,
        max_cols,
        previous_table=previous_table,
        previous_sheet_name=previous_sheet_name,
    )

    _style_check_area(worksheet, table, formula_cells, max_cols)
    _apply_output_widths(worksheet, table, max_cols)

    if formula_cells:
        note = "Có công thức kiểm tra tổng đơn giản. Cần rà soát lại nếu bảng có cấu trúc nhóm phức tạp."
        result = TableCheckResult(table.index, "Có kiểm tra", len(formula_cells), 0, note, formula_cells)
        _write_sheet_legend(worksheet)
        return _finalize_result(worksheet, max_cols, result)

    if table.statement_type in {StatementType.GENERAL_INFO, StatementType.SIGNATURE}:
        _write_sheet_legend(worksheet)
        result = TableCheckResult(table.index, "Không kiểm tra", 0, 0, f"{table.statement_type.value}.")
        return _finalize_result(worksheet, max_cols, result)

    if (
        table.statement_type == StatementType.NOTE
        and _is_non_arithmetic_information_table(table)
    ) or (
        table.statement_type == StatementType.UNKNOWN
        and _rows_with_accounting_values(table) == 0
    ):
        _write_sheet_legend(worksheet)
        result = TableCheckResult(
            table.index,
            "Không kiểm tra",
            0,
            0,
            (
                "Bảng cung cấp thông tin mô tả, không có cấu trúc tổng để kiểm tra số học."
                if table.statement_type == StatementType.NOTE
                else "Bảng cung cấp thông tin mô tả, không có số liệu kế toán để kiểm tra."
            ),
        )
        return _finalize_result(worksheet, max_cols, result)

    _write_sheet_legend(worksheet)
    result = TableCheckResult(
        table.index,
        "Cần xem xét",
        0,
        0,
        "Chưa nhận diện được cấu trúc tổng phù hợp để tự động gắn công thức.",
    )
    return _finalize_result(worksheet, max_cols, result)


def _find_total_rows(table: ExtractedTable) -> list[tuple[int, list[int]]]:
    total_rows: list[tuple[int, list[int]]] = []
    for row_idx, row in enumerate(table.rows):
        label = normalized_key(" ".join(row[:2]))
        if not _looks_like_total_label(label) and not _looks_like_formatted_subtotal(
            table,
            row_idx,
        ):
            continue
        numeric_cols = [
            col_idx
            for col_idx, value in enumerate(row)
            if parse_accounting_number(value) is not None or _is_dash_zero(value)
        ]
        if numeric_cols:
            total_rows.append((row_idx, numeric_cols))
    return total_rows


def _is_narrative_contract_table(table: ExtractedTable) -> bool:
    all_text = normalized_key(" ".join(value for row in table.rows for value in row))
    narrative_markers = (
        "so hop dong",
        "ngay hop dong",
        "ngay dao han",
        "phu luc hop dong",
        "thoi han",
        "muc dich",
        "lai suat",
        "han muc cho vay",
        "tai san the chap",
        "tai san dam bao",
    )
    matches = sum(1 for marker in narrative_markers if marker in all_text)
    has_explicit_total = any(
        "tong cong" in normalized_key(" ".join(row[:2]))
        for row in table.rows
    )
    return matches >= 3 and not has_explicit_total


def _is_non_arithmetic_information_table(table: ExtractedTable) -> bool:
    return (
        _is_narrative_contract_table(table)
        or _is_ownership_percentage_table(table)
        or _is_state_tax_balance_table_without_total(table)
        or _is_depreciation_life_table(table)
        or _is_fully_depreciated_asset_disclosure(table)
        or _is_collateral_disclosure_table(table)
        or _is_single_row_period_table(table)
        or _is_related_party_relationship_table(table)
        or _is_related_party_non_total_table(table)
    )


def _is_ownership_percentage_table(table: ExtractedTable) -> bool:
    text = normalized_key(
        " ".join(
            [table.title_hint, *table.context_hints]
            + [value for row in table.rows for value in row]
        )
    )
    has_percentage_header = any(
        any(marker in normalized_key(value) for marker in ("ti le", "ty le"))
        and any(marker in normalized_key(value) for marker in ("von gop", "%"))
        for row in table.rows[:3]
        for value in row
    )
    has_ownership_context = any(
        marker in text
        for marker in ("chu dau tu", "co dong", "thanh vien gop von", "chu so huu")
    )
    has_total = any(
        _looks_like_total_label(normalized_key(_cell_value(row, 0)))
        for row in table.rows
    )
    return has_percentage_header and has_ownership_context and not has_total


def _is_state_tax_balance_table_without_total(table: ExtractedTable) -> bool:
    labels = {normalized_key(_cell_value(row, 0)) for row in table.rows}
    required_labels = {
        "thue va cac khoan khac phai thu nha nuoc",
        "thue va cac khoan phai nop nha nuoc",
    }
    has_total = any(
        _looks_like_total_label(normalized_key(_cell_value(row, 0)))
        for row in table.rows
    )
    return required_labels.issubset(labels) and not has_total


def _is_contract_schedule_table(table: ExtractedTable) -> bool:
    header = normalized_key(
        " ".join(value for row in table.rows[:3] for value in row)
    )
    markers = (
        "so hop dong",
        "ngay hop dong",
        "ngay dao han",
        "thoi han",
        "muc dich",
        "lai suat",
        "tai san the chap",
        "tai san dam bao",
        "han muc cho vay",
    )
    matches = sum(1 for marker in markers if marker in header)
    return matches >= 3


def _is_related_party_relationship_table(table: ExtractedTable) -> bool:
    context = normalized_key(" ".join((table.title_hint, *table.context_hints)))
    header = normalized_key(
        " ".join(value for row in table.rows[:3] for value in row)
    )
    return (
        "giao dich voi cac ben lien quan" in context
        and "ben lien quan" in header
        and "moi quan he" in header
        and _rows_with_accounting_values(table) == 0
    )


def _is_related_party_non_total_table(table: ExtractedTable) -> bool:
    if table.statement_type != StatementType.NOTE:
        return False
    context = normalized_key(" ".join((table.title_hint, *table.context_hints)))
    header = normalized_key(
        " ".join(value for row in table.rows[:3] for value in row)
    )
    has_period_columns = (
        ("nam nay" in header and "nam truoc" in header)
        or ("so cuoi nam" in header and "so dau nam" in header)
    )
    has_total = any(
        _looks_like_total_label(normalized_key(_cell_value(row, 0)))
        for row in table.rows
    )
    return (
        "ben lien quan" in context
        and has_period_columns
        and _rows_with_accounting_values(table) > 0
        and not has_total
    )


def _is_future_lease_commitment_table(table: ExtractedTable) -> bool:
    text = normalized_key(
        " ".join(
            [table.title_hint, *table.context_hints]
            + [value for row in table.rows for value in row]
        )
    )
    return all(
        marker in text
        for marker in ("tien thue", "phai tra trong tuong lai", "hop dong thue")
    )


def _is_inventory_provision_rollforward(table: ExtractedTable) -> bool:
    text = normalized_key(
        " ".join([table.title_hint, *table.context_hints] + [value for row in table.rows for value in row])
    )
    is_provision = "du phong" in text or "du phong giam gia" in text
    has_rollforward = "trich lap" in text or "hoan nhap" in text or ("so dau nam" in text and "so cuoi nam" in text)
    return is_provision and has_rollforward


def _is_non_reconcilable_note_table(table: ExtractedTable) -> bool:
    return (
        is_tax_rollforward_note(table)
        or _is_inventory_provision_rollforward(table)
        or _is_narrative_contract_table(table)
        or _is_contract_schedule_table(table)
        or _is_depreciation_life_table(table)
        or _is_fully_depreciated_asset_disclosure(table)
        or _is_related_party_relationship_table(table)
        or _is_related_party_non_total_table(table)
        or _is_future_lease_commitment_table(table)
        or _is_single_row_off_balance_table(table)
    )


def _is_single_row_off_balance_table(table: ExtractedTable) -> bool:
    text = normalized_key(
        " ".join(
            [table.title_hint, *table.context_hints]
            + [value for row in table.rows for value in row]
        )
    )
    has_period_pair = (
        ("so cuoi nam" in text and "so dau nam" in text)
        or ("nam nay" in text and "nam truoc" in text)
    )
    has_off_balance_marker = any(
        marker in text
        for marker in (
            "ngoai bang can doi",
            "ngoai te cac loai",
            "cac khoan muc ngoai bang",
        )
    )
    return (
        has_period_pair
        and has_off_balance_marker
        and _rows_with_accounting_values(table) == 1
    )


def _is_depreciation_life_table(table: ExtractedTable) -> bool:
    text = normalized_key(
        " ".join(
            [table.title_hint, *table.context_hints]
            + [value for row in table.rows for value in row]
        )
    )
    if "so nam khau hao" in text or "thoi gian khau hao" in text:
        return True

    has_year_header = any(
        normalized_key(value) == "nam"
        for row in table.rows[:3]
        for value in row
    )
    range_count = sum(
        1
        for row in table.rows
        for value in row[1:]
        if re.fullmatch(r"\d+\s*[-–—]\s*\d+", clean_text(value))
    )
    asset_class_count = sum(
        1
        for row in table.rows
        if any(
            marker in normalized_key(_cell_value(row, 0))
            for marker in (
                "nha cua",
                "may moc",
                "phuong tien van tai",
                "thiet bi van phong",
            )
        )
    )
    return has_year_header and range_count >= 2 and asset_class_count >= 2


def _is_fully_depreciated_asset_disclosure(table: ExtractedTable) -> bool:
    text = normalized_key(
        " ".join(
            [table.title_hint, *table.context_hints]
            + [value for row in table.rows for value in row]
        )
    )
    return "khau hao het" in text and any(
        marker in text for marker in ("van con su dung", "dang con su dung")
    )


def _is_collateral_disclosure_table(table: ExtractedTable) -> bool:
    text = normalized_key(" ".join(value for row in table.rows for value in row))
    has_collateral_marker = any(
        marker in text
        for marker in ("the chap", "cam co", "dam bao cac khoan vay", "tai san dam bao")
    )
    return has_collateral_marker and _rows_with_accounting_values(table) <= 1


def _is_single_row_period_table(table: ExtractedTable) -> bool:
    header_text = normalized_key(
        " ".join(value for row in table.rows[:3] for value in row)
    )
    has_current_period = "so cuoi nam" in header_text or "nam nay" in header_text
    has_prior_period = "so dau nam" in header_text or "nam truoc" in header_text
    return (
        has_current_period
        and has_prior_period
        and _rows_with_accounting_values(table) == 1
    )


def _rows_with_accounting_values(table: ExtractedTable) -> int:
    return sum(
        1
        for row in table.rows
        if any(
            parse_accounting_number(value) is not None or _is_dash_zero(value)
            for value in row[1:]
        )
    )


def _cell_value(row: list[str], column: int) -> str:
    return row[column] if column < len(row) else ""


def _write_note_table_checks(
    worksheet,
    table: ExtractedTable,
    max_cols: int,
    *,
    previous_table: ExtractedTable | None = None,
    previous_sheet_name: str = "",
) -> list[str]:
    formula_cells: list[str] = []
    if (
        not _is_receivable_net_continuation(table, previous_table)
        and not _is_income_tax_schedule(table)
    ):
        formula_cells.extend(_write_vertical_note_checks(worksheet, table, max_cols))
    formula_cells.extend(_write_horizontal_note_checks(worksheet, table, max_cols))
    formula_cells.extend(_write_tax_rollforward_checks(worksheet, table, max_cols))
    formula_cells.extend(_write_income_tax_schedule_checks(worksheet, table))
    formula_cells.extend(
        _write_receivable_net_continuation_checks(
            worksheet,
            table,
            previous_table,
            previous_sheet_name,
            max_cols,
        )
    )
    return formula_cells


def _write_vertical_note_checks(worksheet, table: ExtractedTable, max_cols: int) -> list[str]:
    check_rows = _find_total_rows(table)
    if not check_rows:
        return []

    prepared_checks: list[
        tuple[int, list[tuple[int, str, tuple[int, int]]]]
    ] = []
    for target_row, numeric_cols in check_rows:
        column_checks: list[tuple[int, str, tuple[int, int]]] = []
        for numeric_col in numeric_cols:
            fixed_asset_components = _fixed_asset_net_value_components(
                table,
                target_row,
                numeric_col,
            )
            if fixed_asset_components is not None:
                column_checks.append(
                    (numeric_col, "fixed_asset_net", fixed_asset_components)
                )
                continue
            net_components = _net_revenue_components(table, target_row, numeric_col)
            if net_components is not None:
                column_checks.append((numeric_col, "net_revenue", net_components))
                continue
            deduction_components = _deduction_total_components(
                table,
                target_row,
                numeric_col,
            )
            if deduction_components is not None:
                column_checks.append(
                    (numeric_col, "deduction_net", deduction_components)
                )
                continue
            sum_range = _vertical_sum_range(table, target_row, numeric_col)
            if sum_range is not None:
                column_checks.append((numeric_col, "sum", sum_range))
        if column_checks:
            prepared_checks.append((target_row, column_checks))

    if not prepared_checks:
        return []

    formula_cells: list[str] = []
    output_row = DATA_START_ROW + table.row_count + 1

    for target_row, column_checks in prepared_checks:
        target_excel_row = target_row + DATA_START_ROW
        label_cell = worksheet.cell(
            row=output_row,
            column=1,
            value=f"Kiểm tra cộng dọc dòng {target_excel_row}",
        )
        label_cell.border = THIN_BORDER
        for numeric_col, check_kind, source_rows in column_checks:
            excel_col = numeric_col + 1
            col_letter = get_column_letter(excel_col)
            start_row, end_row = source_rows
            if check_kind in {"net_revenue", "deduction_net", "fixed_asset_net"}:
                expression = (
                    f"{col_letter}{target_excel_row}"
                    f"-{col_letter}{start_row + DATA_START_ROW}"
                    f"{'+' if check_kind in {'net_revenue', 'deduction_net'} else '-'}"
                    f"{col_letter}{end_row + DATA_START_ROW}"
                )
            else:
                expression = (
                    f"{col_letter}{target_excel_row}"
                    f"-SUM({col_letter}{start_row + DATA_START_ROW}:"
                    f"{col_letter}{end_row + DATA_START_ROW})"
                )
            formula = _excel_check_formula(expression)
            diff_cell = worksheet.cell(row=output_row, column=excel_col, value=formula)
            diff_cell.number_format = CHECK_DIFFERENCE_NUMBER_FORMAT
            diff_cell.border = THIN_BORDER
            formula_cells.append(diff_cell.coordinate)
            if check_kind in {"net_revenue", "deduction_net", "fixed_asset_net"}:
                second_coefficient = (
                    1 if check_kind in {"net_revenue", "deduction_net"} else -1
                )
                direct_diff = (
                    _numeric_cell_value(worksheet, target_excel_row, excel_col)
                    - _numeric_cell_value(
                        worksheet,
                        start_row + DATA_START_ROW,
                        excel_col,
                    )
                    + second_coefficient * _numeric_cell_value(
                        worksheet,
                        end_row + DATA_START_ROW,
                        excel_col,
                    )
                )
            else:
                direct_diff = _numeric_cell_value(
                    worksheet,
                    target_excel_row,
                    excel_col,
                ) - _sum_numeric_range(
                    worksheet,
                    start_row + DATA_START_ROW,
                    end_row + DATA_START_ROW,
                    excel_col,
                )
            if _has_check_difference(direct_diff):
                _mark_arithmetic_difference(
                    worksheet.cell(row=target_excel_row, column=excel_col), direct_diff
                )
        output_row += 1

    return formula_cells


def _write_horizontal_note_checks(worksheet, table: ExtractedTable, max_cols: int) -> list[str]:
    total_cols = _horizontal_total_columns(table)
    if not total_cols:
        return []

    formula_cells: list[str] = []
    numeric_columns = _numeric_columns(table)
    check_col = max_cols + 1
    header = worksheet.cell(row=HEADER_ROW, column=check_col, value="Kiểm tra cộng ngang")
    header.font = Font(bold=True)
    header.border = THIN_BORDER
    worksheet.column_dimensions[get_column_letter(check_col)].width = 20

    for row_idx, row in enumerate(table.rows):
        excel_row = row_idx + DATA_START_ROW
        row_formulas: list[str] = []
        row_has_issue = False
        for total_col in total_cols:
            if _is_merged_continuation(table, row_idx, total_col - 1):
                continue
            if len(row) < total_col:
                continue
            total_value = parse_accounting_number(row[total_col - 1])
            if total_value is None:
                continue
            detail_cols = [
                col_idx
                for col_idx in range(1, total_col)
                if len(row) >= col_idx
                and not _is_merged_continuation(table, row_idx, col_idx - 1)
                and _is_horizontal_numeric_value(row[col_idx - 1], col_idx, numeric_columns)
            ]
            if len(detail_cols) < 2:
                continue
            total_letter = get_column_letter(total_col)
            detail_refs = ",".join(f"{get_column_letter(col_idx)}{excel_row}" for col_idx in detail_cols)
            row_formulas.append(f"{total_letter}{excel_row}-SUM({detail_refs})")
            direct_diff = _numeric_cell_value(worksheet, excel_row, total_col) - sum(
                _numeric_cell_value(worksheet, excel_row, col_idx) for col_idx in detail_cols
            )
            if _has_check_difference(direct_diff):
                row_has_issue = True
                _mark_arithmetic_difference(
                    worksheet.cell(row=excel_row, column=total_col), direct_diff
                )
        if not row_formulas:
            continue
        formula = _excel_check_formula("+".join(f"({part})" for part in row_formulas))
        diff_cell = worksheet.cell(row=excel_row, column=check_col, value=formula)
        diff_cell.number_format = CHECK_DIFFERENCE_NUMBER_FORMAT
        diff_cell.border = THIN_BORDER
        formula_cells.append(diff_cell.coordinate)

    return formula_cells


def _horizontal_total_columns(table: ExtractedTable) -> list[int]:
    total_cols: list[int] = []
    layout = _table_layout(table)
    numeric_columns = _numeric_columns(table)
    for col_idx in range(layout.code_col + 1, table.column_count + 1):
        if any(
            _is_merged_continuation(table, row_idx, col_idx - 1)
            for row_idx in range(min(3, len(table.rows)))
        ):
            continue
        header_text = " ".join(
            row[col_idx - 1]
            for row in table.rows[:3]
            if len(row) >= col_idx and row[col_idx - 1]
        )
        key = normalized_key(header_text)
        has_checkable_rows = any(
            len(row) >= col_idx
            and parse_accounting_number(row[col_idx - 1]) is not None
            and sum(
                1
                for detail_col in range(1, col_idx)
                if len(row) >= detail_col
                and _is_horizontal_numeric_value(row[detail_col - 1], detail_col, numeric_columns)
            ) >= 2
            for row in table.rows
        )
        if has_checkable_rows and ("tong cong" in key or key == "tong" or key.endswith(" tong")):
            total_cols.append(col_idx)
    return total_cols


def _is_merged_continuation(
    table: ExtractedTable,
    row_idx: int,
    col_idx: int,
) -> bool:
    """Nhận diện ô phụ của vùng merge theo tọa độ zero-based trong dữ liệu nguồn."""
    return any(
        start_row <= row_idx <= end_row
        and start_col <= col_idx <= end_col
        and (row_idx, col_idx) != (start_row, start_col)
        for start_row, start_col, end_row, end_col in table.merged_ranges
    )


def _write_tax_rollforward_checks(
    worksheet,
    table: ExtractedTable,
    max_cols: int,
) -> list[str]:
    columns = _tax_rollforward_columns(table)
    if columns is None:
        return []
    opening_col, payable_col, paid_col, closing_col = columns
    check_col = max_cols + 1
    header = worksheet.cell(
        row=HEADER_ROW,
        column=check_col,
        value="Kiểm tra cộng ngang thuế",
    )
    header.font = Font(bold=True)
    header.border = THIN_BORDER
    worksheet.column_dimensions[get_column_letter(check_col)].width = 28

    formula_cells: list[str] = []
    for row_index, row in enumerate(table.rows):
        values = [_tax_numeric_value(_cell_value(row, col - 1)) for col in columns]
        if any(value is None for value in values):
            continue
        opening, payable, paid, closing = values
        excel_row = DATA_START_ROW + row_index
        opening_ref = f"{get_column_letter(opening_col)}{excel_row}"
        payable_ref = f"{get_column_letter(payable_col)}{excel_row}"
        paid_ref = f"{get_column_letter(paid_col)}{excel_row}"
        closing_ref = f"{get_column_letter(closing_col)}{excel_row}"
        formula = _excel_check_formula(
            f"{closing_ref}-{opening_ref}-{payable_ref}+ABS({paid_ref})"
        )
        diff_cell = worksheet.cell(row=excel_row, column=check_col, value=formula)
        diff_cell.number_format = CHECK_DIFFERENCE_NUMBER_FORMAT
        diff_cell.border = THIN_BORDER
        formula_cells.append(diff_cell.coordinate)
        direct_diff = closing - opening - payable + abs(paid)
        if _has_check_difference(direct_diff):
            _mark_arithmetic_difference(
                worksheet.cell(row=excel_row, column=closing_col), direct_diff
            )
    return formula_cells


def _tax_rollforward_columns(
    table: ExtractedTable,
) -> tuple[int, int, int, int] | None:
    headers = {
        col: normalized_key(
            " ".join(_cell_value(row, col - 1) for row in table.rows[:3])
        )
        for col in range(1, table.column_count + 1)
    }
    opening_col = next(
        (col for col, header in headers.items() if "so dau" in header or "01/01" in header),
        None,
    )
    payable_col = next(
        (col for col, header in headers.items() if "so phai nop" in header),
        None,
    )
    paid_col = next(
        (
            col
            for col, header in headers.items()
            if "so da nop" in header and "khau tru" in header
        ),
        None,
    )
    closing_col = next(
        (col for col, header in headers.items() if "so cuoi" in header or "31/12" in header),
        None,
    )
    if None in {opening_col, payable_col, paid_col, closing_col}:
        return None
    columns = (
        int(opening_col),
        int(payable_col),
        int(paid_col),
        int(closing_col),
    )
    if len(set(columns)) != 4:
        return None
    return columns


def _tax_numeric_value(value: object) -> Decimal | None:
    if _is_dash_zero(value):
        return Decimal(0)
    return parse_accounting_number(value)


def _is_income_tax_schedule(table: ExtractedTable) -> bool:
    labels = {normalized_key(_cell_value(row, 0)) for row in table.rows}
    return {
        "loi nhuan truoc thue",
        "thu nhap chiu thue",
        "thu nhap tinh thue",
    }.issubset(labels)


def _write_income_tax_schedule_checks(worksheet, table: ExtractedTable) -> list[str]:
    labels = [normalized_key(_cell_value(row, 0)) for row in table.rows]
    row_by_label = {
        label: index
        for index, label in enumerate(labels)
        if label
    }
    profit_row = row_by_label.get("loi nhuan truoc thue")
    increase_row = row_by_label.get("dieu chinh tang loi nhuan truoc thue")
    decrease_row = row_by_label.get("dieu chinh giam loi nhuan truoc thue")
    taxable_rows = [index for index, label in enumerate(labels) if label == "thu nhap chiu thue"]
    loss_row = row_by_label.get("so chuyen lo mang sang")
    assessable_row = row_by_label.get("thu nhap tinh thue")
    expense_row = row_by_label.get("chi phi thue thu nhap doanh nghiep hien hanh")
    if not taxable_rows or profit_row is None:
        return []

    rules: list[tuple[str, int, list[tuple[int, str]]]] = []
    taxable_row = taxable_rows[0]
    if increase_row is not None and decrease_row is not None:
        rules.append(
            (
                "Thu nhập chịu thuế",
                taxable_row,
                [(profit_row, "signed"), (increase_row, "signed"), (decrease_row, "reduction")],
            )
        )
    taxable_details = _income_tax_breakdown_rows(table, taxable_row)
    if taxable_details:
        rules.append(
            (
                "Chi tiết thu nhập chịu thuế",
                taxable_row,
                [(row, "signed") for row in taxable_details],
            )
        )
    if assessable_row is not None and loss_row is not None:
        rules.append(
            (
                "Thu nhập tính thuế",
                assessable_row,
                [(taxable_row, "signed"), (loss_row, "reduction")],
            )
        )
        assessable_details = _income_tax_breakdown_rows(table, assessable_row)
        if assessable_details:
            rules.append(
                (
                    "Chi tiết thu nhập tính thuế",
                    assessable_row,
                    [(row, "signed") for row in assessable_details],
                )
            )
    if expense_row is not None:
        expense_components = [
            (
                index,
                (
                    "reduction"
                    if "thue thu nhap doanh nghiep duoc mien" in label
                    else "signed"
                ),
            )
            for index, label in enumerate(labels[:expense_row])
            if (
                "thue thu nhap doanh nghiep phai nop" in label
                or "dieu chinh chi phi thue thu nhap doanh nghiep" in label
                or "thue thu nhap doanh nghiep duoc mien" in label
                or "thue toi thieu toan cau" in label
            )
        ]
        if expense_components:
            rules.append(
                (
                    "Chi phí thuế TNDN hiện hành",
                    expense_row,
                    expense_components,
                )
            )
    if not rules:
        return []

    period_cols = sorted(_numeric_columns(table))
    if not period_cols:
        return []
    output_row = worksheet.max_row + 2
    title = worksheet.cell(row=output_row, column=1, value="Kiểm tra bảng tính thuế")
    title.font = Font(bold=True)
    title.border = THIN_BORDER
    output_row += 1
    formula_cells: list[str] = []
    for description, target_row, sources in rules:
        label_cell = worksheet.cell(row=output_row, column=1, value=description)
        label_cell.border = THIN_BORDER
        for col in period_cols:
            target_value = _tax_numeric_value(_cell_value(table.rows[target_row], col - 1))
            source_values = [
                _tax_numeric_value(_cell_value(table.rows[source_row], col - 1))
                for source_row, _ in sources
            ]
            if target_value is None or any(value is None for value in source_values):
                continue
            target_excel_row = DATA_START_ROW + target_row
            target_ref = f"{get_column_letter(col)}{target_excel_row}"
            expression = target_ref
            expected = Decimal(0)
            for (source_row, mode), source_value in zip(sources, source_values):
                source_ref = f"{get_column_letter(col)}{DATA_START_ROW + source_row}"
                if mode == "reduction":
                    expression += f"+ABS({source_ref})"
                    expected -= abs(source_value)
                else:
                    expression += f"-{source_ref}"
                    expected += source_value
            diff_cell = worksheet.cell(
                row=output_row,
                column=col,
                value=_excel_check_formula(expression),
            )
            diff_cell.number_format = CHECK_DIFFERENCE_NUMBER_FORMAT
            diff_cell.border = THIN_BORDER
            formula_cells.append(diff_cell.coordinate)
            direct_diff = target_value - expected
            if _has_check_difference(direct_diff):
                _mark_arithmetic_difference(
                    worksheet.cell(row=target_excel_row, column=col), direct_diff
                )
        output_row += 1
    return formula_cells


def _write_receivable_net_continuation_checks(
    worksheet,
    table: ExtractedTable,
    previous_table: ExtractedTable | None,
    previous_sheet_name: str,
    max_cols: int,
) -> list[str]:
    if not _is_receivable_net_continuation(table, previous_table):
        return []
    provision_row = next(
        (
            index
            for index, row in enumerate(table.rows)
            if normalized_key(_cell_value(row, 0)).startswith("du phong phai thu")
        ),
        None,
    )
    net_row = next(
        (
            index
            for index, row in enumerate(table.rows)
            if normalized_key(_cell_value(row, 0)).startswith("gia tri thuan")
        ),
        None,
    )
    previous_total_row = next(
        (
            index
            for index, row in enumerate(previous_table.rows)
            if "tong cong" in normalized_key(" ".join(row[:2]))
        ),
        None,
    )
    if None in {provision_row, net_row, previous_total_row}:
        return []

    provision_groups = _unique_numeric_groups(table.rows[int(provision_row)])
    net_groups = _unique_numeric_groups(table.rows[int(net_row)])
    total_groups = _unique_numeric_groups(previous_table.rows[int(previous_total_row)])
    if min(len(provision_groups), len(net_groups), len(total_groups)) < 2:
        return []

    escaped_sheet = previous_sheet_name.replace("'", "''")
    check_start_col = max_cols + 1
    formula_cells: list[str] = []
    for period_index, period_name in enumerate(("Kỳ này", "Kỳ trước")):
        provision_col, provision_value = provision_groups[period_index]
        net_col, net_value = net_groups[period_index]
        total_col, total_value = total_groups[period_index]
        net_excel_row = DATA_START_ROW + int(net_row)
        provision_excel_row = DATA_START_ROW + int(provision_row)
        previous_total_excel_row = DATA_START_ROW + int(previous_total_row)
        check_col = check_start_col + period_index
        header = worksheet.cell(
            row=HEADER_ROW,
            column=check_col,
            value=f"Giá trị thuần - {period_name}",
        )
        header.font = Font(bold=True)
        header.border = THIN_BORDER
        worksheet.column_dimensions[get_column_letter(check_col)].width = 24
        net_ref = f"{get_column_letter(net_col)}{net_excel_row}"
        provision_ref = f"{get_column_letter(provision_col)}{provision_excel_row}"
        total_ref = (
            f"'{escaped_sheet}'!"
            f"{get_column_letter(total_col)}{previous_total_excel_row}"
        )
        formula = _excel_check_formula(f"{net_ref}-{total_ref}-{provision_ref}")
        diff_cell = worksheet.cell(row=net_excel_row, column=check_col, value=formula)
        diff_cell.number_format = CHECK_DIFFERENCE_NUMBER_FORMAT
        diff_cell.border = THIN_BORDER
        formula_cells.append(diff_cell.coordinate)
        direct_diff = net_value - total_value - provision_value
        if _has_check_difference(direct_diff):
            _mark_arithmetic_difference(
                worksheet.cell(row=net_excel_row, column=net_col), direct_diff
            )
    return formula_cells


def _is_receivable_net_continuation(
    table: ExtractedTable,
    previous_table: ExtractedTable | None,
) -> bool:
    if (
        previous_table is None
        or table.statement_type != StatementType.NOTE
        or previous_table.statement_type != StatementType.NOTE
        or normalized_key(table.title_hint) != normalized_key(previous_table.title_hint)
    ):
        return False
    labels = [normalized_key(_cell_value(row, 0)) for row in table.rows]
    return (
        any(label.startswith("du phong phai thu") for label in labels)
        and any(label.startswith("gia tri thuan") for label in labels)
    )


def _unique_numeric_groups(row: list[str]) -> list[tuple[int, Decimal]]:
    groups: list[tuple[int, Decimal]] = []
    previous_col = -2
    previous_value: Decimal | None = None
    for col_index, raw_value in enumerate(row, start=1):
        value = _tax_numeric_value(raw_value)
        if value is None:
            previous_col = -2
            previous_value = None
            continue
        if col_index == previous_col + 1 and value == previous_value:
            previous_col = col_index
            continue
        groups.append((col_index, value))
        previous_col = col_index
        previous_value = value
    return groups


def _income_tax_breakdown_rows(table: ExtractedTable, target_row: int) -> list[int]:
    section_row = next(
        (
            index
            for index in range(target_row + 1, min(len(table.rows), target_row + 3))
            if normalized_key(_cell_value(table.rows[index], 0)).startswith("trong do")
        ),
        None,
    )
    if section_row is None:
        return []
    details: list[int] = []
    for index in range(section_row + 1, len(table.rows)):
        if _is_blank_row(table.rows[index]):
            break
        label = normalized_key(_cell_value(table.rows[index], 0))
        if not label:
            break
        if label.startswith("trong do") or label in {"so chuyen lo mang sang", "thu nhap tinh thue"}:
            break
        details.append(index)
    return details


def _is_dash_zero(value: object) -> bool:
    return clean_text(value) in DASH_ZERO_MARKERS


def _expected_number_style(table: ExtractedTable) -> str | None:
    counts: dict[str, int] = {}
    for row in table.rows:
        for value in row:
            if parse_accounting_number(value) is None:
                continue
            style = _number_separator_style(value)
            if style is None:
                continue
            counts[style] = counts.get(style, 0) + 1
    if not counts:
        return None
    style, count = max(counts.items(), key=lambda item: item[1])
    return style if count >= 3 else None


def _has_number_format_issue(value: object, expected_style: str | None) -> bool:
    if expected_style is None:
        return False
    actual_style = _number_separator_style(value)
    if actual_style is None or actual_style == "plain":
        return False
    return (
        expected_style == "dot_thousand"
        and actual_style == "comma_thousand"
    ) or (
        expected_style == "comma_thousand"
        and actual_style == "dot_thousand"
    )


def _number_separator_style(value: object) -> str | None:
    text = clean_text(value)
    if not text or text in DASH_ZERO_MARKERS:
        return None
    if parse_accounting_number(text) is None:
        return None

    text = text.replace("VND", "").replace("VNĐ", "").replace("vnđ", "")
    text = text.replace("USD", "").strip()
    text = text.strip("()")
    text = text.lstrip("-").replace(" ", "")
    if not re.search(r"\d", text):
        return None

    has_dot = "." in text
    has_comma = "," in text
    if has_dot and has_comma:
        return "decimal_comma" if text.rfind(",") > text.rfind(".") else "decimal_dot"
    if has_dot:
        parts = text.split(".")
        if len(parts) > 1 and all(len(part) == 3 for part in parts[1:]):
            return "dot_thousand"
        return "decimal_dot"
    if has_comma:
        parts = text.split(",")
        if len(parts) > 1 and all(len(part) == 3 for part in parts[1:]):
            return "comma_thousand"
        return "decimal_comma"
    return "plain"


def _apply_output_widths(worksheet, table: ExtractedTable, max_cols: int) -> None:
    if table.statement_type == StatementType.NOTE:
        for col_idx, width_inches in enumerate(table.column_widths[:max_cols], start=1):
            if width_inches <= 0:
                continue
            worksheet.column_dimensions[get_column_letter(col_idx)].width = min(
                52,
                max(8, width_inches * 10),
            )
    largest_numeric_width = max(
        (
            _numeric_display_width(cell.value)
            for row in worksheet.iter_rows()
            for cell in row
            if isinstance(cell.value, (int, float))
        ),
        default=0,
    )
    for col_idx in range(1, max(max_cols + 3, worksheet.max_column) + 1):
        column_letter = get_column_letter(col_idx)
        cells = [
            worksheet.cell(row=row_idx, column=col_idx)
            for row_idx in range(1, worksheet.max_row + 1)
        ]
        numeric_width = max(
            (_numeric_display_width(cell.value) for cell in cells if isinstance(cell.value, (int, float))),
            default=0,
        )
        has_formula = any(
            isinstance(cell.value, str) and cell.value.startswith("=")
            for cell in cells
        )
        required_width = max(
            18,
            numeric_width + 2,
            largest_numeric_width + 2 if has_formula else 0,
        )
        worksheet.column_dimensions[column_letter].width = max(
            worksheet.column_dimensions[column_letter].width or 0,
            required_width,
        )
    if max_cols and table.statement_type != StatementType.NOTE:
        worksheet.column_dimensions["A"].width = max(
            worksheet.column_dimensions["A"].width or 0,
            38,
        )


def _numeric_display_width(value: int | float) -> int:
    integer_value = int(round(abs(float(value))))
    return len(f"{integer_value:,}") + (2 if value < 0 else 0)


def _apply_note_source_presentation(worksheet, table: ExtractedTable) -> None:
    for (row_idx, col_idx), presentation in table.cell_presentations.items():
        cell = worksheet.cell(row=row_idx + DATA_START_ROW, column=col_idx + 1)
        cell.font = Font(
            bold=presentation.bold,
            italic=presentation.italic,
            underline="single" if presentation.underline else None,
            color=presentation.font_color or None,
        )
        if presentation.fill_color:
            cell.fill = PatternFill("solid", fgColor=presentation.fill_color)
        cell.alignment = Alignment(
            horizontal=presentation.horizontal_alignment or None,
            vertical=presentation.vertical_alignment or None,
            wrap_text=True,
        )
        cell.border = Border(
            top=_presentation_side(presentation.top_border, cell.border.top),
            right=_presentation_side(presentation.right_border, cell.border.right),
            bottom=_presentation_side(presentation.bottom_border, cell.border.bottom),
            left=_presentation_side(presentation.left_border, cell.border.left),
        )

    _highlight_total_rows(worksheet, table)

    for start_row, start_col, end_row, end_col in table.merged_ranges:
        worksheet.merge_cells(
            start_row=start_row + DATA_START_ROW,
            start_column=start_col + 1,
            end_row=end_row + DATA_START_ROW,
            end_column=end_col + 1,
        )


def _highlight_total_rows(worksheet, table: ExtractedTable) -> None:
    for source_row_idx, row in enumerate(table.rows):
        if not _looks_like_total_label(normalized_key(" ".join(row[:2]))):
            continue
        excel_row = source_row_idx + DATA_START_ROW
        for col_idx in range(1, table.column_count + 1):
            cell = worksheet.cell(row=excel_row, column=col_idx)
            cell.font = Font(
                name=cell.font.name,
                size=cell.font.size,
                bold=True,
                italic=cell.font.italic,
                underline=cell.font.underline,
                color=cell.font.color,
            )


def _presentation_side(presentation_border, fallback: Side) -> Side:
    if not presentation_border.style:
        return fallback
    return Side(style=presentation_border.style, color=presentation_border.color or "000000")


def _style_check_area(
    worksheet,
    table: ExtractedTable,
    formula_cells: list[str],
    max_cols: int,
) -> None:
    for coordinate in formula_cells:
        worksheet[coordinate].fill = LINKED_VALUE_FILL
    for col_idx in range(max_cols + 1, worksheet.max_column + 1):
        header = worksheet.cell(row=HEADER_ROW, column=col_idx)
        if header.value:
            header.fill = HEADER_FILL
            header.font = Font(bold=True, color=COLOR_WHITE)
    first_check_row = DATA_START_ROW + table.row_count + 1
    for row_idx in range(first_check_row, worksheet.max_row + 1):
        label_cell = worksheet.cell(row=row_idx, column=1)
        if isinstance(label_cell.value, str) and label_cell.value.startswith("Kiểm tra"):
            label_cell.fill = GROUP_HEADER_FILL if label_cell.font.bold else CHECK_VALUE_FILL


def _finalize_result(worksheet, max_cols: int, result: TableCheckResult) -> TableCheckResult:
    issue_count = _count_direct_issue_cells(worksheet, max_cols)
    return TableCheckResult(
        table_index=result.table_index,
        status=result.status,
        check_count=result.check_count,
        issue_count=issue_count,
        note=result.note,
        formula_cells=result.formula_cells,
    )


def _count_direct_issue_cells(worksheet, max_cols: int = 0) -> int:
    count = 0
    col_limit = max_cols if max_cols > 0 else worksheet.max_column
    for row_idx in range(DATA_START_ROW, worksheet.max_row + 1):
        for col_idx in range(1, col_limit + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            if cell.fill and cell.fill.fgColor and COLOR_YELLOW in str(cell.fill.fgColor.rgb):
                count += 1
    return count


def _write_sheet_legend(worksheet) -> None:
    pass


def _numeric_cell_value(worksheet, row_idx: int, col_idx: int) -> float:
    value = worksheet.cell(row=row_idx, column=col_idx).value
    return float(value) if isinstance(value, (int, float)) else 0.0


def _sum_numeric_range(worksheet, row_start: int, row_end: int, col_idx: int) -> float:
    return sum(_numeric_cell_value(worksheet, row_idx, col_idx) for row_idx in range(row_start, row_end + 1))


def _add_single_issue_highlight(
    worksheet,
    row_idx: int,
    value_col: int,
    check_col: int,
    has_direct_issue: bool = False,
    direct_difference: Decimal | float | int | None = None,
) -> None:
    value_letter = get_column_letter(value_col)
    check_letter = get_column_letter(check_col)
    worksheet.conditional_formatting.add(
        f"{value_letter}{row_idx}",
        FormulaRule(formula=[f"${check_letter}{row_idx}<>0"], fill=WARNING_FILL),
    )
    if has_direct_issue:
        cell = worksheet.cell(row=row_idx, column=value_col)
        if direct_difference is None:
            cell.fill = WARNING_FILL
        else:
            _mark_arithmetic_difference(cell, direct_difference)


def _add_pair_issue_highlight(
    worksheet,
    row_idx: int,
    layout: TableLayout,
    check_col: int,
    has_current_issue: bool = False,
    has_prior_issue: bool = False,
    current_difference: Decimal | float | int | None = None,
    prior_difference: Decimal | float | int | None = None,
) -> None:
    _add_single_issue_highlight(
        worksheet, row_idx, layout.current_col, check_col, has_current_issue, current_difference
    )
    _add_single_issue_highlight(
        worksheet, row_idx, layout.prior_col, check_col + 1, has_prior_issue, prior_difference
    )


def _calculate_arithmetic_diff(
    current_type: StatementType,
    current_layout: TableLayout,
    rule: ArithmeticRule,
    period: str,
    current_values: dict[str, dict[int, float]],
    primary_values: dict[StatementType, dict[str, dict[int, float]]],
    primary_layouts: dict[StatementType, TableLayout],
) -> float:
    value_col = _period_col(current_type, current_type, period, current_layout, primary_layouts)
    target = current_values.get(_normalize_code(rule.target_code), {}).get(value_col, 0.0)
    total = 0.0
    for term in rule.terms:
        source_type = term.source or current_type
        source_col = _period_col(current_type, source_type, period, current_layout, primary_layouts)
        if source_type == current_type:
            value = current_values.get(_normalize_code(term.code), {}).get(source_col, 0.0)
        else:
            value = _lookup_value(primary_values, source_type, term.code, source_col)
        total += term.coefficient * value
    return target - total


def _calculate_sign_diff(
    rule: SignRule,
    value_col: int,
    values: dict[str, dict[int, float]],
) -> float:
    value = values.get(_normalize_code(rule.code), {}).get(value_col, 0.0)
    if rule.expected == "positive":
        return value if value < 0 else 0.0
    return value if value > 0 else 0.0


def _calculate_cross_period_diff(
    current_type: StatementType,
    rule: CrossPeriodRule,
    primary_values: dict[StatementType, dict[str, dict[int, float]]],
    primary_layouts: dict[StatementType, TableLayout],
) -> float:
    target_col = _period_col(current_type, current_type, rule.target_period, primary_layouts[current_type], primary_layouts)
    source_col = _period_col(current_type, rule.source, rule.source_period, primary_layouts[current_type], primary_layouts)
    target = _lookup_value(primary_values, current_type, rule.target_code, target_col)
    source = _lookup_value(primary_values, rule.source, rule.source_code, source_col)
    return target - source


def _lookup_value(
    primary_values: dict[StatementType, dict[str, dict[int, float]]],
    statement_type: StatementType,
    code: str,
    value_col: int,
) -> float:
    return primary_values.get(statement_type, {}).get(_normalize_code(code), {}).get(value_col, 0.0)


def _write_business_checks(
    worksheet,
    table: ExtractedTable,
    layout: TableLayout,
    primary_sheets: dict[StatementType, str],
    primary_layouts: dict[StatementType, TableLayout],
    primary_code_rows: dict[StatementType, dict[str, int]],
    primary_values: dict[StatementType, dict[str, dict[int, float]]],
    check_start_col: int,
    rule_pack: RulePack,
) -> TableCheckResult:
    rules = rule_pack.arithmetic_rules.get(table.statement_type, ())
    sign_rules = rule_pack.sign_rules.get(table.statement_type, ())
    cross_period_rules = rule_pack.cross_period_rules.get(table.statement_type, ())
    if not rules and not sign_rules and not cross_period_rules:
        return TableCheckResult(table.index, "Không có rule nghiệp vụ")

    code_rows = _code_row_lookup(table, data_start_row=DATA_START_ROW)
    if not code_rows:
        return TableCheckResult(
            table.index,
            "Cần xem xét",
            0,
            0,
            "Không xác định được cột mã số để áp dụng rule nghiệp vụ.",
        )

    current_values = _table_value_lookup(table, layout)
    current_typed_values = table_value_state_lookup(table, layout)
    formula_cells: list[str] = []
    slot_by_row: dict[int, int] = {}
    needs_data_review = False

    for rule in rules:
        needs_data_review = _write_inline_arithmetic_check(
            worksheet,
            table,
            layout,
            code_rows,
            primary_sheets,
            primary_layouts,
            primary_code_rows,
            primary_values,
            current_values,
            current_typed_values,
            rule,
            check_start_col,
            slot_by_row,
            formula_cells,
        ) or needs_data_review

    for rule in sign_rules:
        _write_inline_sign_check(
            worksheet,
            code_rows,
            rule,
            layout,
            current_values,
            check_start_col,
            slot_by_row,
            formula_cells,
        )

    for rule in cross_period_rules:
        _write_inline_cross_period_check(
            worksheet,
            table,
            layout,
            code_rows,
            primary_sheets,
            primary_layouts,
            primary_code_rows,
            primary_values,
            rule,
            check_start_col,
            slot_by_row,
            formula_cells,
        )

    if formula_cells or needs_data_review:
        return TableCheckResult(
            table.index,
            "Cần xem xét" if needs_data_review else "Có rule nghiệp vụ",
            len(formula_cells),
            0,
            (
                "Có dữ liệu thiếu hoặc không hợp lệ trong rule số học; cần xác minh."
                if needs_data_review
                else "Đã gắn công thức kiểm tra theo mã số chỉ tiêu."
            ),
            formula_cells,
        )
    return TableCheckResult(table.index, "Cần xem xét", 0, 0, "Có rule nhưng thiếu mã số cần thiết trong bảng.")


def _write_inline_arithmetic_check(
    worksheet,
    table: ExtractedTable,
    layout: TableLayout,
    code_rows: dict[str, int],
    primary_sheets: dict[StatementType, str],
    primary_layouts: dict[StatementType, TableLayout],
    primary_code_rows: dict[StatementType, dict[str, int]],
    primary_values: dict[StatementType, dict[str, dict[int, float]]],
    current_values: dict[str, dict[int, float]],
    current_typed_values,
    rule: ArithmeticRule,
    check_start_col: int,
    slot_by_row: dict[int, int],
    formula_cells: list[str],
) -> bool:
    target_row = code_rows.get(_normalize_code(rule.target_code))
    if target_row is None:
        component_codes = {
            _normalize_code(term.code)
            for term in rule.terms
            if term.source is None and _normalize_code(term.code) in code_rows
        }
        if not component_codes:
            return False
        message = (
            "Kiểm tra số học: Có mã thành phần nhưng thiếu mã đích "
            f"{rule.target_code}."
        )
        has_actionable_component = False
        for component_code in component_codes:
            row = code_rows[component_code]
            for value_column in {layout.current_col, layout.prior_col}:
                typed_value = current_typed_values.get(component_code, {}).get(value_column)
                if typed_value is not None and typed_value.status in {
                    TableValueStatus.ZERO,
                    TableValueStatus.DASH_ZERO,
                }:
                    continue
                cell = worksheet.cell(row=row, column=value_column)
                cell.fill = WARNING_FILL
                if cell.comment is None:
                    cell.comment = Comment(message, "CHECK_FS_RULE")
                has_actionable_component = True
        return has_actionable_component

    if all(term.source is None for term in rule.terms) and not any(
        _normalize_code(term.code) in code_rows for term in rule.terms
    ):
        return False

    slot = _reserve_check_slot(worksheet, target_row, check_start_col, layout, slot_by_row)
    note = _rule_description(rule)
    missing_codes: set[str] = set()
    review_messages: list[str] = []

    for period, check_col in (("current", slot), ("prior", slot + 1)):
        value_col = _period_col(table.statement_type, table.statement_type, period, layout, primary_layouts)
        target_ref = _cell_ref(table.statement_type, table.statement_type, rule.target_code, value_col, code_rows, primary_sheets, primary_code_rows)
        if target_ref is None:
            continue
        terms: list[str] = []
        for term in rule.terms:
            source_type = term.source or table.statement_type
            source_col = _period_col(table.statement_type, source_type, period, layout, primary_layouts)
            ref = _cell_ref(table.statement_type, source_type, term.code, source_col, code_rows, primary_sheets, primary_code_rows)
            if ref is None:
                missing_codes.add(term.code)
                continue
            if term.coefficient == 1:
                terms.append(ref)
            elif term.coefficient == -1:
                terms.append(f"-{ref}")
            else:
                terms.append(f"{term.coefficient}*{ref}")
        same_table_rule = all(term.source is None for term in rule.terms)
        typed_result = None
        if same_table_rule:
            typed_result = calculate_same_table_arithmetic(
                rule,
                value_col,
                current_typed_values,
            )
        expression = "+".join(terms).replace("+-", "-") if terms else "0"
        formula = _excel_check_formula(f"{target_ref}-({expression})")
        if typed_result is not None and not typed_result.applicable:
            continue
        cell_value = 0 if typed_result is not None and typed_result.needs_review else formula
        diff_cell = worksheet.cell(row=target_row, column=check_col, value=cell_value)
        diff_cell.number_format = CHECK_DIFFERENCE_NUMBER_FORMAT
        diff_cell.border = THIN_BORDER
        if typed_result is not None and typed_result.needs_review:
            period_label = "kỳ này" if period == "current" else "kỳ trước"
            review_messages.append(
                f"{period_label}: mã {', '.join(typed_result.review_codes)} thiếu hoặc không hợp lệ"
            )
        else:
            formula_cells.append(diff_cell.coordinate)

    if all(term.source is None for term in rule.terms):
        current_result = calculate_same_table_arithmetic(rule, layout.current_col, current_typed_values)
        prior_result = calculate_same_table_arithmetic(rule, layout.prior_col, current_typed_values)
        current_diff = current_result.difference
        prior_diff = prior_result.difference
        current_issue = current_result.needs_review or (
            current_result.applicable
            and current_diff is not None
            and _has_check_difference(current_diff)
        )
        prior_issue = prior_result.needs_review or (
            prior_result.applicable
            and prior_diff is not None
            and _has_check_difference(prior_diff)
        )
    else:
        current_diff = _calculate_arithmetic_diff(table.statement_type, layout, rule, "current", current_values, primary_values, primary_layouts)
        prior_diff = _calculate_arithmetic_diff(table.statement_type, layout, rule, "prior", current_values, primary_values, primary_layouts)
        current_issue = _has_check_difference(current_diff)
        prior_issue = _has_check_difference(prior_diff)
    _add_pair_issue_highlight(
        worksheet,
        target_row,
        layout,
        slot,
        current_issue,
        prior_issue,
        current_diff,
        prior_diff,
    )
    _write_inline_note(
        worksheet,
        target_row,
        slot,
        note,
        sorted(missing_codes),
        review_messages,
    )
    return bool(review_messages)


def _write_inline_sign_check(
    worksheet,
    code_rows: dict[str, int],
    rule: SignRule,
    layout: TableLayout,
    current_values: dict[str, dict[int, float]],
    check_start_col: int,
    slot_by_row: dict[int, int],
    formula_cells: list[str],
) -> None:
    row = code_rows.get(rule.code)
    if row is None:
        return
    slot = _reserve_check_slot(worksheet, row, check_start_col, layout, slot_by_row)
    for value_col, check_col in ((layout.current_col, slot), (layout.prior_col, slot + 1)):
        value_ref = f"{get_column_letter(value_col)}{row}"
        expression = (
            f"IF({value_ref}<0,{value_ref},0)"
            if rule.expected == "positive"
            else f"IF({value_ref}>0,{value_ref},0)"
        )
        formula = _excel_check_formula(expression)
        diff_cell = worksheet.cell(row=row, column=check_col, value=formula)
        diff_cell.number_format = CHECK_DIFFERENCE_NUMBER_FORMAT
        diff_cell.border = THIN_BORDER
        formula_cells.append(diff_cell.coordinate)
    current_diff = _calculate_sign_diff(rule, layout.current_col, current_values)
    prior_diff = _calculate_sign_diff(rule, layout.prior_col, current_values)
    _add_pair_issue_highlight(
        worksheet,
        row,
        layout,
        slot,
        _has_check_difference(current_diff),
        _has_check_difference(prior_diff),
        current_diff,
        prior_diff,
    )
    _write_inline_note(worksheet, row, slot, rule.note, [])


def _write_inline_cross_period_check(
    worksheet,
    table: ExtractedTable,
    layout: TableLayout,
    code_rows: dict[str, int],
    primary_sheets: dict[StatementType, str],
    primary_layouts: dict[StatementType, TableLayout],
    primary_code_rows: dict[StatementType, dict[str, int]],
    primary_values: dict[StatementType, dict[str, dict[int, float]]],
    rule: CrossPeriodRule,
    check_start_col: int,
    slot_by_row: dict[int, int],
    formula_cells: list[str],
) -> None:
    target_row = code_rows.get(_normalize_code(rule.target_code))
    if target_row is None:
        return
    target_col = _period_col(table.statement_type, table.statement_type, rule.target_period, layout, primary_layouts)
    source_col = _period_col(table.statement_type, rule.source, rule.source_period, layout, primary_layouts)
    target_ref = _cell_ref(table.statement_type, table.statement_type, rule.target_code, target_col, code_rows, primary_sheets, primary_code_rows)
    source_ref = _cell_ref(table.statement_type, rule.source, rule.source_code, source_col, code_rows, primary_sheets, primary_code_rows)
    if target_ref is None:
        return
    missing = [] if source_ref is not None else [rule.source_code]
    formula = _excel_check_formula(f"{target_ref}-{source_ref or 0}")
    slot = _reserve_check_slot(worksheet, target_row, check_start_col, layout, slot_by_row)
    check_col = slot if target_col == 4 else slot + 1
    diff_cell = worksheet.cell(row=target_row, column=check_col, value=formula)
    diff_cell.number_format = CHECK_DIFFERENCE_NUMBER_FORMAT
    diff_cell.border = THIN_BORDER
    formula_cells.append(diff_cell.coordinate)
    value_col = target_col
    direct_diff = _calculate_cross_period_diff(table.statement_type, rule, primary_values, primary_layouts)
    _add_single_issue_highlight(
        worksheet,
        target_row,
        value_col,
        check_col,
        _has_check_difference(direct_diff),
        direct_diff,
    )
    _write_inline_note(worksheet, target_row, slot, rule.note, missing)


def _write_inline_check_headers(
    worksheet,
    check_start_col: int,
    slot_index: int,
    layout: TableLayout,
) -> None:
    base_col = check_start_col + slot_index * 3
    current_letter = get_column_letter(layout.current_col)
    prior_letter = get_column_letter(layout.prior_col)
    headers = (f"Kiểm tra cột {current_letter}", f"Kiểm tra cột {prior_letter}", "Ghi chú")
    if slot_index:
        headers = (
            f"Kiểm tra {current_letter} ({slot_index + 1})",
            f"Kiểm tra {prior_letter} ({slot_index + 1})",
            f"Ghi chú ({slot_index + 1})",
        )
    for offset, header in enumerate(headers):
        cell = worksheet.cell(row=HEADER_ROW, column=base_col + offset, value=header)
        cell.font = Font(bold=True)
        cell.border = THIN_BORDER
        worksheet.column_dimensions[get_column_letter(base_col + offset)].width = 18 if offset < 2 else 52


def _reserve_check_slot(
    worksheet,
    row_idx: int,
    check_start_col: int,
    layout: TableLayout,
    slot_by_row: dict[int, int],
) -> int:
    slot_index = slot_by_row.get(row_idx, 0)
    slot_by_row[row_idx] = slot_index + 1
    _write_inline_check_headers(worksheet, check_start_col, slot_index, layout)
    return check_start_col + slot_index * 3


def _write_inline_note(
    worksheet,
    row_idx: int,
    check_col: int,
    note: str,
    missing_codes: list[str],
    review_messages: list[str] | None = None,
) -> None:
    note_text = note
    if missing_codes:
        note_text = f"{note}. Thiếu mã: {', '.join(missing_codes)}"
    if review_messages:
        note_text = f"Cần xác minh dữ liệu ({'; '.join(review_messages)}). {note}"
    left = get_column_letter(check_col)
    right = get_column_letter(check_col + 1)
    note_cell = worksheet.cell(
        row=row_idx,
        column=check_col + 2,
        value=(
            note_text
            if review_messages
            else f'=IF(OR({left}{row_idx}<>0,{right}{row_idx}<>0),"{note_text}","")'
        ),
    )
    note_cell.border = THIN_BORDER


def _rule_description(rule: ArithmeticRule) -> str:
    if rule.note != "Khác số chi tiết":
        return rule.note
    terms = []
    for term in rule.terms:
        prefix = "-" if term.coefficient == -1 else ""
        terms.append(f"{prefix}{term.code}")
    return f"{rule.target_code} khác tổng chi tiết ({' + '.join(terms)})"


def _nearest_statement_tables(
    tables: list[ExtractedTable],
    current_table: ExtractedTable,
) -> dict[StatementType, ExtractedTable]:
    statement_types = {
        StatementType.BALANCE_SHEET_ASSETS,
        StatementType.BALANCE_SHEET_EQUITY,
        StatementType.INCOME_STATEMENT,
        StatementType.CASH_FLOW,
    }
    peers: dict[StatementType, ExtractedTable] = {}
    for statement_type in statement_types:
        candidates = [table for table in tables if table.statement_type == statement_type]
        if not candidates:
            continue
        peers[statement_type] = min(
            candidates,
            key=lambda table: (abs(table.index - current_table.index), table.index),
        )
    return peers


def _primary_sheet_lookup(
    tables: list[ExtractedTable],
    sheet_names: dict[int, str],
) -> dict[StatementType, str]:
    lookup: dict[StatementType, str] = {}
    for table in tables:
        if table.statement_type not in lookup:
            lookup[table.statement_type] = sheet_names[table.index]
    return lookup


def _primary_code_row_lookup(tables: list[ExtractedTable]) -> dict[StatementType, dict[str, int]]:
    lookup: dict[StatementType, dict[str, int]] = {}
    for table in tables:
        if table.statement_type not in lookup:
            lookup[table.statement_type] = _code_row_lookup(
                table,
                _table_layout(table),
                data_start_row=DATA_START_ROW,
            )
    return lookup


def _primary_value_lookup(tables: list[ExtractedTable]) -> dict[StatementType, dict[str, dict[int, float]]]:
    lookup: dict[StatementType, dict[str, dict[int, float]]] = {}
    for table in tables:
        if table.statement_type in lookup:
            continue
        lookup[table.statement_type] = _table_value_lookup(table, _table_layout(table))
    return lookup


def _primary_layout_lookup(tables: list[ExtractedTable]) -> dict[StatementType, TableLayout]:
    lookup: dict[StatementType, TableLayout] = {}
    for table in tables:
        if table.statement_type not in lookup:
            lookup[table.statement_type] = _table_layout(table)
    return lookup


def _cell_ref(
    current_type: StatementType,
    source_type: StatementType,
    code: str,
    value_col: int,
    current_code_rows: dict[str, int],
    primary_sheets: dict[StatementType, str],
    primary_code_rows: dict[StatementType, dict[str, int]],
) -> str | None:
    if source_type == current_type:
        row = current_code_rows.get(_normalize_code(code))
        sheet_name = None
    else:
        row = primary_code_rows.get(source_type, {}).get(_normalize_code(code))
        sheet_name = primary_sheets.get(source_type)
    if row is None:
        return None
    cell = f"{get_column_letter(value_col)}{row}"
    if sheet_name:
        return f"'{sheet_name}'!{cell}"
    return cell


def _vertical_sum_range(table: ExtractedTable, target_row_idx: int, col_idx: int) -> tuple[int, int] | None:
    end_idx = target_row_idx - 1
    while end_idx >= 0 and _is_blank_row(table.rows[end_idx]):
        end_idx -= 1
    if end_idx < 0:
        return None

    start_idx = end_idx
    target_label = normalized_key(" ".join(table.rows[target_row_idx][:2]))
    crosses_blank_group_separators = "tong cong" in target_label
    is_closing_balance = _looks_like_closing_balance_label(target_label)
    while start_idx > 0:
        previous_row = table.rows[start_idx - 1]
        if _is_blank_row(previous_row):
            if crosses_blank_group_separators:
                start_idx -= 1
                continue
            break
        previous_label = normalized_key(" ".join(previous_row[:2]))
        if is_closing_balance and _looks_like_opening_balance_label(previous_label):
            start_idx -= 1
            break
        if _looks_like_total_row(previous_row):
            break
        if (
            not crosses_blank_group_separators
            and _looks_like_formatted_subtotal(table, start_idx - 1)
        ):
            break
        start_idx -= 1

    while start_idx <= end_idx:
        row = table.rows[start_idx]
        if col_idx < len(row) and _is_vertical_numeric_value(row[col_idx]):
            break
        start_idx += 1

    numeric_count = 0
    for idx in range(start_idx, end_idx + 1):
        row = table.rows[idx]
        if col_idx < len(row) and _is_vertical_numeric_value(row[col_idx]):
            numeric_count += 1

    target_has_explicit_total_label = _looks_like_total_label(target_label)
    minimum_numeric_count = (
        1
        if target_has_explicit_total_label
        or _looks_like_formatted_subtotal(table, target_row_idx)
        else 2
    )
    if numeric_count < minimum_numeric_count:
        return None
    return start_idx, end_idx


def _is_vertical_numeric_value(value: object) -> bool:
    return parse_accounting_number(value) is not None or _is_dash_zero(value)


def _is_horizontal_numeric_value(value: object, col_idx: int, numeric_columns: set[int]) -> bool:
    return parse_accounting_number(value) is not None or (_is_dash_zero(value) and col_idx in numeric_columns)


def _is_blank_row(row: list[str]) -> bool:
    return not any(clean_text(value) for value in row)


def _looks_like_total_row(row: list[str]) -> bool:
    label = normalized_key(" ".join(row[:2]))
    return _looks_like_total_label(label)


def _looks_like_total_label(label: str) -> bool:
    return (
        "tong cong" in label
        or "tong tai san" in label
        or "tong nguon von" in label
        or label.startswith("gia tri thuan")
        or (label.startswith("gia tri con lai") and " cua " not in f" {label} ")
        or label.startswith("doanh thu thuan")
        or label.startswith("so du cuoi")
        or label.startswith("so cuoi")
        or label.startswith("so dau")
        or label.startswith("cong cuoi")
        or label.startswith("cong dau")
    )


def _looks_like_closing_balance_label(label: str) -> bool:
    return label.startswith("so cuoi") or label.startswith("so du cuoi")


def _looks_like_opening_balance_label(label: str) -> bool:
    is_regular_opening = label.startswith("so dau") or label.startswith("so du dau")
    is_transition_balance = _looks_like_closing_balance_label(label) and (
        "dau nam" in label or "dau ky" in label
    )
    return is_regular_opening or is_transition_balance


def _looks_like_formatted_subtotal(table: ExtractedTable, row_idx: int) -> bool:
    row = table.rows[row_idx]
    if not row or clean_text(row[0]):
        return False
    numeric_columns = [
        col_idx
        for col_idx, value in enumerate(row)
        if parse_accounting_number(value) is not None or _is_dash_zero(value)
    ]
    if not numeric_columns:
        return False
    for col_idx in numeric_columns:
        presentation = table.cell_presentations.get((row_idx, col_idx))
        if presentation is None:
            continue
        border_styles = {
            presentation.bottom_border.style,
            presentation.top_border.style,
        }
        border_styles.discard("")
        if presentation.bold and border_styles:
            return True
        if border_styles.intersection({"double", "thick"}):
            return True
    return False


def _net_revenue_components(
    table: ExtractedTable,
    target_row_idx: int,
    col_idx: int,
) -> tuple[int, int] | None:
    target_label = normalized_key(" ".join(table.rows[target_row_idx][:2]))
    if not target_label.startswith("doanh thu thuan"):
        return None

    subtotal_rows: list[int] = []
    for row_idx in range(target_row_idx - 1, -1, -1):
        if _looks_like_formatted_subtotal(table, row_idx):
            row = table.rows[row_idx]
            if col_idx < len(row) and _is_vertical_numeric_value(row[col_idx]):
                subtotal_rows.append(row_idx)
                if len(subtotal_rows) == 2:
                    break
    if len(subtotal_rows) < 2:
        deduction_header = next(
            (
                row_idx
                for row_idx in range(target_row_idx - 1, -1, -1)
                if normalized_key(_cell_value(table.rows[row_idx], 0)).startswith("tru")
            ),
            None,
        )
        if deduction_header is None:
            return None
        deduction_row = next(
            (
                row_idx
                for row_idx in range(deduction_header + 1, target_row_idx)
                if col_idx < len(table.rows[row_idx])
                and _is_vertical_numeric_value(table.rows[row_idx][col_idx])
            ),
            None,
        )
        base_row = next(
            (
                row_idx
                for row_idx in range(deduction_header - 1, -1, -1)
                if col_idx < len(table.rows[row_idx])
                and _is_vertical_numeric_value(table.rows[row_idx][col_idx])
            ),
            None,
        )
        if base_row is None or deduction_row is None:
            return None
        return base_row, deduction_row

    deduction_row, base_row = subtotal_rows
    has_deduction_header = any(
        normalized_key(table.rows[row_idx][0]).startswith("tru")
        for row_idx in range(base_row + 1, deduction_row)
        if table.rows[row_idx]
    )
    if not has_deduction_header:
        return None
    return base_row, deduction_row


def _deduction_total_components(
    table: ExtractedTable,
    target_row_idx: int,
    col_idx: int,
) -> tuple[int, int] | None:
    target_label = normalized_key(" ".join(table.rows[target_row_idx][:2]))
    if "tong cong" not in target_label:
        return None

    subtotal_rows: list[int] = []
    for row_idx in range(target_row_idx - 1, -1, -1):
        if not _looks_like_formatted_subtotal(table, row_idx):
            continue
        row = table.rows[row_idx]
        if col_idx < len(row) and _is_vertical_numeric_value(row[col_idx]):
            subtotal_rows.append(row_idx)
            if len(subtotal_rows) == 2:
                break
    if len(subtotal_rows) < 2:
        return None

    deduction_row, base_row = subtotal_rows
    has_deduction_header = any(
        normalized_key(_cell_value(table.rows[row_idx], 0)).startswith("tru")
        for row_idx in range(base_row + 1, deduction_row)
    )
    if not has_deduction_header:
        return None
    return base_row, deduction_row


def _fixed_asset_net_value_components(
    table: ExtractedTable,
    target_row_idx: int,
    col_idx: int,
) -> tuple[int, int] | None:
    target_label = normalized_key(" ".join(table.rows[target_row_idx][:2]))
    if target_label.startswith("gia tri con lai"):
        compact_components = _compact_fixed_asset_net_components(
            table,
            target_row_idx,
            col_idx,
        )
        if compact_components is not None:
            return compact_components
    target_period = _balance_period_key(target_label)
    if target_period is None:
        return None

    net_section_row = _nearest_section_row(
        table,
        target_row_idx,
        "gia tri con lai",
    )
    if net_section_row is None:
        return None
    depreciation_section_row = _nearest_section_row(
        table,
        net_section_row,
        ("gia tri hao mon", "gia tri khau hao"),
    )
    if depreciation_section_row is None:
        return None
    cost_section_row = _nearest_section_row(
        table,
        depreciation_section_row,
        "nguyen gia",
    )
    if cost_section_row is None:
        return None

    cost_row = _find_period_row_in_section(
        table,
        cost_section_row + 1,
        depreciation_section_row,
        target_period,
        col_idx,
    )
    depreciation_row = _find_period_row_in_section(
        table,
        depreciation_section_row + 1,
        net_section_row,
        target_period,
        col_idx,
    )
    if cost_row is None or depreciation_row is None:
        return None
    return cost_row, depreciation_row


def _compact_fixed_asset_net_components(
    table: ExtractedTable,
    target_row_idx: int,
    col_idx: int,
) -> tuple[int, int] | None:
    cost_row: int | None = None
    depreciation_row: int | None = None
    for row_idx in range(target_row_idx - 1, -1, -1):
        label = normalized_key(_cell_value(table.rows[row_idx], 0))
        if depreciation_row is None and label.startswith(
            ("khau hao luy ke", "hao mon luy ke", "gia tri khau hao", "gia tri hao mon")
        ):
            depreciation_row = row_idx
            continue
        if cost_row is None and label.startswith("nguyen gia"):
            cost_row = row_idx
        if cost_row is not None and depreciation_row is not None:
            break
    if cost_row is None or depreciation_row is None:
        return None
    if not all(
        col_idx < len(table.rows[row_idx])
        and _is_vertical_numeric_value(table.rows[row_idx][col_idx])
        for row_idx in (cost_row, depreciation_row)
    ):
        return None
    return cost_row, depreciation_row


def _nearest_section_row(
    table: ExtractedTable,
    before_row: int,
    section_label: str | tuple[str, ...],
) -> int | None:
    section_labels = (section_label,) if isinstance(section_label, str) else section_label
    for row_idx in range(before_row - 1, -1, -1):
        label = normalized_key(" ".join(table.rows[row_idx][:2]))
        if label.startswith(section_labels):
            return row_idx
    return None


def _balance_period_key(label: str) -> str | None:
    if label.startswith("so dau") or label.startswith("so du dau"):
        return "opening"
    if label.startswith("so cuoi") or label.startswith("so du cuoi"):
        return "closing"
    return None


def _find_period_row_in_section(
    table: ExtractedTable,
    start_row: int,
    end_row: int,
    period: str,
    col_idx: int,
) -> int | None:
    for row_idx in range(start_row, end_row):
        row = table.rows[row_idx]
        label = normalized_key(" ".join(row[:2]))
        if _balance_period_key(label) != period:
            continue
        if col_idx < len(row) and _is_vertical_numeric_value(row[col_idx]):
            return row_idx
    return None


def _write_summary(
    summary,
    tables: list[ExtractedTable],
    results: list[TableCheckResult],
    note_matches_by_table: dict[int, list[NoteMatchResult]],
    rule_pack: RulePack,
    *,
    maturity_parent_indexes: set[int] | None = None,
) -> None:
    maturity_parent_indexes = maturity_parent_indexes or set()
    headers = [
        "STT bảng",
        "Sheet chi tiết",
        "Nội dung bảng",
        "Loại bảng",
        "Dòng",
        "Cột",
        "Số phép kiểm tra",
        "Tổng sai sót",
        "Trạng thái",
        "Trạng thái đối chiếu BCTC",
        "Ghi chú",
    ]
    for col_idx, header in enumerate(headers, start=1):
        cell = summary.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True, color=COLOR_WHITE)
        cell.fill = TITLE_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER

    by_index = {result.table_index: result for result in results}
    for row_idx, table in enumerate(tables, start=2):
        result = by_index[table.index]
        sheet_name = _sheet_name(table)
        values = [
            table.index,
            sheet_name,
            table_content_label(table),
            table_type_label(table),
            table.row_count,
            table.column_count,
            result.check_count,
            result.issue_count,
            result.status,
            _note_reconciliation_summary_status(
                table,
                note_matches_by_table.get(table.index, []),
                maturity_split_parent=table.index in maturity_parent_indexes,
            ),
            result.note,
        ]
        for col_idx, value in enumerate(values, start=1):
            cell = summary.cell(row=row_idx, column=col_idx, value=value)
            cell.border = THIN_BORDER
            if col_idx == 2:
                cell.hyperlink = f"#{sheet_name}!A1"
                cell.style = "Hyperlink"
        status_cell = summary.cell(row=row_idx, column=9)
        status_cell.fill = (
            TABLE_STATUS_FILLS["Cần xem xét"]
            if result.status == "Cần xem xét"
            else WARNING_FILL
            if result.issue_count
            else TABLE_STATUS_FILLS.get(result.status, CHECK_VALUE_FILL)
        )
        reconciliation_status_cell = summary.cell(row=row_idx, column=10)
        reconciliation_status_cell.fill = _note_reconciliation_summary_fill(
            reconciliation_status_cell.value
        )
        if result.issue_count:
            summary.cell(row=row_idx, column=8).fill = WARNING_FILL
        else:
            summary.cell(row=row_idx, column=8).fill = CHECK_VALUE_FILL

    widths = [12, 18, 48, 38, 10, 10, 18, 14, 16, 28, 70]
    for col_idx, width in enumerate(widths, start=1):
        summary.column_dimensions[get_column_letter(col_idx)].width = width
    summary.freeze_panes = "A2"
    summary.auto_filter.ref = f"A1:K{max(1, summary.max_row)}"
    summary["M1"] = "Bộ rule nghiệp vụ"
    summary["M1"].font = Font(bold=True, color=COLOR_WHITE)
    summary["M1"].fill = TITLE_FILL
    summary["N1"] = rule_pack.display_name
    summary["M2"] = "Phiên bản rule"
    summary["M2"].font = Font(bold=True)
    summary["N2"] = rule_pack.version
    summary["M3"] = "Phím tắt Excel"
    summary["M3"].font = Font(bold=True)
    summary["N3"] = "Ctrl + PgDn (Tiếp) | Ctrl + PgUp (Trước)"
    summary.column_dimensions["M"].width = 22
    summary.column_dimensions["N"].width = 38


def _note_reconciliation_summary_status(
    table: ExtractedTable,
    matches: list[NoteMatchResult],
    *,
    maturity_split_parent: bool = False,
) -> str:
    if table.statement_type != StatementType.NOTE:
        return "Không áp dụng"
    if not matches:
        return (
            "Không áp dụng"
            if maturity_split_parent or _is_non_reconcilable_note_table(table)
            else "Chưa có kết quả"
        )

    statuses = {match.status for match in matches}
    for status in ("Difference", "Note value not found", "Statement not found"):
        if status in statuses:
            return _note_status_label(status)
    if statuses == {"Matched"}:
        return _note_status_label("Matched")
    return "Cần xác minh"


def _note_reconciliation_summary_fill(status: str):
    if status == "Khớp":
        return CHECK_VALUE_FILL
    if status == "Có chênh lệch":
        return WARNING_FILL
    if status in {
        "Không tìm thấy số liệu TM",
        "Không tìm thấy chỉ tiêu BCTC",
        "Chưa có kết quả",
        "Cần xác minh",
    }:
        return WARNING_FILL
    return CHECK_VALUE_FILL


def _update_summary_backlinks(
    worksheets_by_index: dict[int, object],
    summary_tables: list[ExtractedTable],
) -> None:
    for summary_row, table in enumerate(summary_tables, start=2):
        worksheet = worksheets_by_index[table.index]
        back_cell = worksheet.cell(row=INFO_ROW, column=2, value="Trở về Tổng hợp")
        back_cell.hyperlink = f"#00_Tong_hop!B{summary_row}"
        back_cell.font = Font(bold=True, color=COLOR_HYPERLINK, underline="single")
        back_cell.alignment = Alignment(horizontal="left", vertical="center")


def _write_detail_navigation(
    worksheets_by_index: dict[int, object],
    detail_tables: list[ExtractedTable],
) -> None:
    """Tạo thanh điều hướng tuần tự theo layout mới (A1: Trước, B1: Tiêu đề, A2: Tiếp, B2: Trở về Tổng hợp)."""
    for position, table in enumerate(detail_tables):
        worksheet = worksheets_by_index[table.index]
        previous_table = detail_tables[position - 1] if position else None
        next_table = detail_tables[position + 1] if position + 1 < len(detail_tables) else None

        for r in (TITLE_ROW, INFO_ROW):
            for col in range(3, max(6, worksheet.max_column) + 1):
                c = worksheet.cell(row=r, column=col)
                if c.value in {"Trở về Tổng hợp", "← Trước", "Tiếp →"}:
                    c.value = None
                    c.hyperlink = None
                    c.border = Border()
                    c.fill = PatternFill()

        prev_cell = worksheet.cell(row=TITLE_ROW, column=1, value="← Trước")
        prev_cell.fill = PatternFill()
        prev_cell.border = Border()
        prev_cell.alignment = Alignment(horizontal="center", vertical="center")
        if previous_table is not None:
            prev_cell.hyperlink = f"#{_sheet_name(previous_table)}!A1"
            prev_cell.font = Font(bold=True, color=COLOR_HYPERLINK, underline="single")
        else:
            prev_cell.hyperlink = None
            prev_cell.font = Font(color=COLOR_MUTED_TEXT, italic=True)

        next_cell = worksheet.cell(row=INFO_ROW, column=1, value="Tiếp →")
        next_cell.fill = PatternFill()
        next_cell.border = Border()
        next_cell.alignment = Alignment(horizontal="center", vertical="center")
        if next_table is not None:
            next_cell.hyperlink = f"#{_sheet_name(next_table)}!A1"
            next_cell.font = Font(bold=True, color=COLOR_HYPERLINK, underline="single")
        else:
            next_cell.hyperlink = None
            next_cell.font = Font(color=COLOR_MUTED_TEXT, italic=True)

        title_cell = worksheet.cell(row=TITLE_ROW, column=2, value=table_content_label(table))
        title_cell.fill = TITLE_FILL
        title_cell.font = Font(bold=True, color=COLOR_WHITE, size=12)
        title_cell.border = THIN_BORDER
        title_cell.alignment = Alignment(horizontal="left", vertical="center")

        home_cell = worksheet.cell(row=INFO_ROW, column=2, value="Trở về Tổng hợp")
        home_cell.fill = PatternFill()
        home_cell.border = Border()
        home_cell.alignment = Alignment(horizontal="left", vertical="center")
        home_cell.font = Font(bold=True, color=COLOR_HYPERLINK, underline="single")
        if not home_cell.hyperlink:
            home_cell.hyperlink = "#00_Tong_hop!A1"


def _write_note_reconciliation_detail(
    worksheet,
    matches: list[NoteMatchResult],
    sheet_names: dict[int, str],
    table: ExtractedTable,
    *,
    maturity_split_parent: bool = False,
) -> None:
    """Ghi kết quả đối chiếu thuộc riêng bảng TM ngay dưới bảng chi tiết."""
    start_row = worksheet.max_row + 2
    last_col = max(1, table.column_count)
    worksheet.merge_cells(
        start_row=start_row,
        start_column=1,
        end_row=start_row,
        end_column=last_col,
    )
    title_cell = worksheet.cell(row=start_row, column=1, value="ĐỐI CHIẾU VỚI BCTC")
    title_cell.font = Font(bold=True, color=COLOR_WHITE, size=12)
    title_cell.fill = TITLE_FILL
    title_cell.border = THIN_BORDER
    title_cell.alignment = Alignment(vertical="center")
    worksheet.row_dimensions[start_row].height = 26

    if not matches:
        message = (
            "Không áp dụng đối chiếu trực tiếp; số liệu được đối chiếu theo bảng phân loại "
            "ngắn hạn/dài hạn bên dưới."
            if maturity_split_parent
            else
            "Không áp dụng đối chiếu BCTC cho bảng này."
            if _is_non_reconcilable_note_table(table)
            else "Chưa có kết quả đối chiếu cho bảng Thuyết minh này."
        )
        message_cell = worksheet.cell(
            row=start_row + 1,
            column=1,
            value=message,
        )
        message_cell.font = Font(italic=True, color=COLOR_MUTED_TEXT)
        message_cell.border = THIN_BORDER
        return

    current_row = start_row + 1
    for match_index, match in enumerate(matches, start=1):
        current_col, prior_col = _note_match_period_columns(table, match)
        note_current_reference = _merged_anchor_reference(table, match.note_current_cell)
        note_prior_reference = _merged_anchor_reference(table, match.note_prior_cell)
        value_start_col = min(current_col, prior_col)
        value_end_col = max(value_start_col, last_col)
        if len(matches) > 1:
            _write_note_detail_section_header(
                worksheet,
                current_row,
                last_col,
                f"Đối chiếu {match_index}: {match.item_name or 'Cần xác minh'}",
            )
            current_row += 1

        statement_sheet_name = sheet_names.get(match.statement_table_index, "")
        bctc_row = current_row
        _write_note_detail_label(worksheet, bctc_row, "Giá trị BCTC")
        for col_idx, cell_reference in (
            (current_col, match.statement_current_cell),
            (prior_col, match.statement_prior_cell),
        ):
            value_cell = worksheet.cell(
                row=bctc_row,
                column=col_idx,
                value=_reconciliation_formula(statement_sheet_name, cell_reference),
            )
            value_cell.number_format = ACCOUNTING_NUMBER_FORMAT
            value_cell.fill = LINKED_VALUE_FILL
            value_cell.border = THIN_BORDER
            _add_formula_source_link(value_cell, statement_sheet_name, cell_reference)

        difference_row = bctc_row + 1
        _write_note_detail_label(worksheet, difference_row, "Chênh lệch")
        for col_idx, note_cell_reference in (
            (current_col, note_current_reference),
            (prior_col, note_prior_reference),
        ):
            difference_cell = worksheet.cell(
                row=difference_row,
                column=col_idx,
                value=_note_difference_formula(col_idx, bctc_row, note_cell_reference),
            )
            difference_cell.number_format = CHECK_DIFFERENCE_NUMBER_FORMAT
            difference_cell.border = THIN_BORDER
            _add_formula_source_link(difference_cell, worksheet.title, note_cell_reference)
            if difference_cell.value is not None and match.status == "Difference":
                difference_cell.fill = STATUS_FILLS["Difference"]

        current_row = difference_row + 1
        if match.status not in {"Matched", "Difference"}:
            current_row = _write_note_detail_metadata_row(
                worksheet,
                current_row,
                value_start_col,
                value_end_col,
                "Cần xác minh",
                _note_status_label(match.status),
                fill=WARNING_FILL,
            )
        current_row += 1


def _note_match_period_columns(table: ExtractedTable, match: NoteMatchResult) -> tuple[int, int]:
    layout = _table_layout(table)
    header_current_col, header_prior_col = _note_header_period_columns(table)
    current_reference = _merged_anchor_reference(table, match.note_current_cell)
    prior_reference = _merged_anchor_reference(table, match.note_prior_cell)
    current_col = (
        _column_from_cell_reference(current_reference)
        or header_current_col
        or layout.current_col
    )
    prior_col = (
        _column_from_cell_reference(prior_reference)
        or header_prior_col
        or layout.prior_col
    )
    if current_col == prior_col and current_reference != prior_reference:
        prior_col = current_col + 1 if current_col < table.column_count else max(1, current_col - 1)
    return max(1, current_col), max(1, prior_col)


def _note_header_period_columns(table: ExtractedTable) -> tuple[int | None, int | None]:
    current_cols, prior_cols = detect_period_columns_from_headers(table)
    return (
        current_cols[0] + 1 if current_cols else None,
        prior_cols[0] + 1 if prior_cols else None,
    )


def _column_from_cell_reference(cell_reference: str) -> int | None:
    match = re.match(r"^([A-Z]+)\d+$", cell_reference.upper())
    if match is None:
        return None
    column_number = 0
    for character in match.group(1):
        column_number = column_number * 26 + ord(character) - ord("A") + 1
    return column_number


def _note_difference_formula(
    column_idx: int,
    bctc_row: int,
    note_cell_reference: str,
) -> str | None:
    if not note_cell_reference:
        return None
    bctc_cell = f"{get_column_letter(column_idx)}{bctc_row}"
    return _excel_check_formula(f"N({bctc_cell})-N({note_cell_reference})")


def _write_note_detail_label(worksheet, row_idx: int, label: str) -> None:
    cell = worksheet.cell(row=row_idx, column=1, value=label)
    cell.font = Font(bold=True, color=COLOR_TEXT)
    cell.fill = GROUP_HEADER_FILL
    cell.border = THIN_BORDER


def _write_note_detail_section_header(worksheet, row_idx: int, last_col: int, title: str) -> None:
    worksheet.merge_cells(
        start_row=row_idx,
        start_column=1,
        end_row=row_idx,
        end_column=last_col,
    )
    cell = worksheet.cell(row=row_idx, column=1, value=title)
    cell.font = Font(bold=True, color=COLOR_TEXT)
    cell.fill = GROUP_HEADER_FILL
    cell.border = THIN_BORDER


def _write_note_detail_metadata_row(
    worksheet,
    row_idx: int,
    value_start_col: int,
    value_end_col: int,
    label: str,
    value,
    *,
    fill: PatternFill | None = None,
    hyperlink: str = "",
) -> int:
    _write_note_detail_label(worksheet, row_idx, label)
    if value_end_col > value_start_col:
        worksheet.merge_cells(
            start_row=row_idx,
            start_column=value_start_col,
            end_row=row_idx,
            end_column=value_end_col,
        )
    value_cell = worksheet.cell(row=row_idx, column=value_start_col, value=value)
    value_cell.alignment = Alignment(vertical="center", wrap_text=True)
    value_cell.border = THIN_BORDER
    if fill is not None:
        value_cell.fill = fill
    if hyperlink:
        value_cell.hyperlink = hyperlink
        value_cell.font = Font(color=COLOR_HYPERLINK, underline="single")
    return row_idx + 1


def _ordered_note_matches(matches: list[NoteMatchResult]) -> list[NoteMatchResult]:
    status_priority = {
        "Difference": 0,
        "Note value not found": 1,
        "Statement not found": 2,
        "Matched": 3,
    }

    def sort_key(match: NoteMatchResult) -> tuple[int, int, float, str]:
        differences = [
            abs(float(value))
            for value in (match.current_difference, match.prior_difference)
            if value is not None
        ]
        largest_difference = max(differences, default=0.0)
        return (
            match.note_table_index,
            status_priority.get(match.status, 4),
            -largest_difference,
            match.statement_code,
        )

    return sorted(matches, key=sort_key)


def _note_status_label(status: str) -> str:
    return NOTE_STATUS_LABELS.get(status, status)


def _reconciliation_formula(sheet_name: str, cell_reference: str):
    if not cell_reference:
        return None
    if not sheet_name:
        return f"={cell_reference}"
    escaped_sheet_name = sheet_name.replace("'", "''")
    return f"='{escaped_sheet_name}'!{cell_reference}"


def _add_formula_source_link(cell, sheet_name: str, cell_reference: str) -> None:
    if not sheet_name or not cell_reference or cell.value is None:
        return
    cell.hyperlink = f"#{sheet_name}!{cell_reference}"
    cell.font = Font(color=COLOR_HYPERLINK, underline="single")
