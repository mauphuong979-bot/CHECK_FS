from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, replace
from decimal import Decimal, ROUND_HALF_UP
import re

from src.domain.models import ExtractedTable, NoteMatchResult, StatementType
from src.extraction.table_layout import detect_period_columns_from_headers
from src.extraction.table_values import DEFAULT_EXCEL_DATA_START_ROW
from src.normalization.number_parser import parse_accounting_number
from src.normalization.text_cleaner import clean_text, normalized_key


DEFAULT_ROUNDING_TOLERANCE = Decimal(0)
RECONCILIATION_ROUND_QUANTUM = Decimal("0.01")
STATEMENT_TYPES = {
    StatementType.BALANCE_SHEET_ASSETS,
    StatementType.BALANCE_SHEET_EQUITY,
    StatementType.INCOME_STATEMENT,
    StatementType.CASH_FLOW,
}
DASH_ZERO_MARKERS = {"-", "–", "—"}


@dataclass(frozen=True)
class StatementItem:
    table_index: int
    source_type: StatementType
    code: str
    note_reference: str
    name: str
    current_value: Decimal | None
    prior_value: Decimal | None
    current_cell: str
    prior_cell: str

    @property
    def source(self) -> str:
        if self.source_type in {
            StatementType.BALANCE_SHEET_ASSETS,
            StatementType.BALANCE_SHEET_EQUITY,
        }:
            return "BS"
        if self.source_type == StatementType.INCOME_STATEMENT:
            return "PL"
        return "CF"


@dataclass(frozen=True)
class AssetNoteCodeRegistry:
    title_markers: tuple[str, ...]
    net_code: str
    cost_code: str
    depreciation_code: str


ASSET_NOTE_CODE_REGISTRIES: tuple[AssetNoteCodeRegistry, ...] = (
    AssetNoteCodeRegistry(("tai san co dinh huu hinh",), "221", "222", "223"),
    AssetNoteCodeRegistry(("tai san co dinh thue tai chinh",), "224", "225", "226"),
    AssetNoteCodeRegistry(("tai san co dinh vo hinh",), "227", "228", "229"),
    AssetNoteCodeRegistry(("bat dong san dau tu",), "230", "231", "232"),
)

NOTE_ITEM_SEMANTIC_MAPPINGS: tuple[tuple[tuple[str, ...], tuple[str, ...]], ...] = (
    (
        ("von chu so huu", "von da gop"),
        ("von dau tu cua chu so huu", "von gop cua chu so huu"),
    ),
)


def reconcile_note_tables(
    tables: list[ExtractedTable],
    rounding_tolerance: Decimal = DEFAULT_ROUNDING_TOLERANCE,
) -> list[NoteMatchResult]:
    """Đối chiếu riêng các bảng Thuyết minh với chỉ tiêu BS/PL/CF."""
    statement_items = _extract_statement_items(tables)
    results: list[NoteMatchResult] = []

    for note_table in tables:
        if note_table.statement_type != StatementType.NOTE:
            continue
        if is_parent_detail_with_maturity_split(note_table, tables):
            continue
        if _is_non_reconcilable_information_note(note_table):
            continue
        if _is_receivable_net_continuation_note(note_table, tables):
            continue
        income_tax_results = _reconcile_income_tax_schedule(
            note_table,
            statement_items,
            rounding_tolerance,
        )
        if income_tax_results is not None:
            results.extend(income_tax_results)
            continue
        depreciation_results = _reconcile_depreciation_expense_to_cash_flow(
            note_table,
            statement_items,
            rounding_tolerance,
        )
        if depreciation_results is not None:
            results.extend(depreciation_results)
            continue
        equity_movement_results = _reconcile_equity_movement_table(
            note_table,
            statement_items,
            rounding_tolerance,
        )
        if equity_movement_results is not None:
            results.extend(equity_movement_results)
            continue
        contributed_capital_results = _reconcile_contributed_capital_table(
            note_table,
            statement_items,
            rounding_tolerance,
        )
        if contributed_capital_results is not None:
            results.extend(contributed_capital_results)
            continue
        state_tax_balance_results = _reconcile_state_tax_balance_rows(
            note_table,
            statement_items,
            rounding_tolerance,
        )
        if state_tax_balance_results is not None:
            results.extend(state_tax_balance_results)
            continue
        asset_results = _reconcile_asset_note_by_codes(
            note_table,
            statement_items,
            rounding_tolerance,
        )
        if asset_results is not None:
            results.extend(asset_results)
            continue
        maturity_results = _reconcile_maturity_rows(
            note_table,
            statement_items,
            rounding_tolerance,
        )
        if maturity_results is not None:
            results.extend(maturity_results)
            continue
        candidates = _find_candidates(note_table, statement_items)
        if not candidates:
            maturity = _title_maturity(note_table)
            current_value, current_cell, prior_value, prior_cell = _note_total_values(note_table)
            has_named_total = bool(_normalized_note_title(note_table)) and _has_reconciliation_total(
                note_table
            )
            results.append(
                NoteMatchResult(
                    note_table_index=note_table.index,
                    note_title=note_table.title_hint,
                    source="BS" if maturity else "",
                    item_name=(
                        _expected_maturity_name(note_table, maturity)
                        if maturity
                        else clean_text(note_table.title_hint)
                        if has_named_total
                        else ""
                    ),
                    note_current=current_value,
                    note_current_cell=current_cell,
                    note_prior=prior_value,
                    note_prior_cell=prior_cell,
                    status="Statement not found",
                    match_type=(
                        "Tiêu đề kỳ hạn và Thuyết minh"
                        if maturity
                        else "Tên bảng + Tổng cộng"
                        if has_named_total
                        else "Tham chiếu Thuyết minh"
                    ),
                )
            )
            continue

        candidate_results = [
            _compare_candidate(note_table, item, match_type, score, rounding_tolerance)
            for item, match_type, score in candidates
        ]
        candidate_results.sort(key=lambda result: (-result.match_score, result.source, result.statement_code))
        results.extend(candidate_results)

    return results


def is_parent_detail_with_maturity_split(
    table: ExtractedTable,
    tables: list[ExtractedTable],
) -> bool:
    """Bỏ đối chiếu bảng chi tiết khi bảng kế tiếp đã phân loại ngắn/dài hạn."""
    try:
        position = tables.index(table)
    except ValueError:
        return False
    if position + 1 >= len(tables):
        return False
    split_table = tables[position + 1]
    if (
        table.statement_type != StatementType.NOTE
        or split_table.statement_type != StatementType.NOTE
        or _maturity_rows(table)
        or len(_maturity_rows(split_table)) != 2
    ):
        return False

    parent_title = _maturity_base_title(_normalized_note_title(table))
    split_title = _maturity_base_title(_normalized_note_title(split_table))
    if not parent_title or parent_title != split_title:
        return False

    parent_current, _, parent_prior, _ = _note_total_values(table)
    split_totals = _maturity_split_totals(split_table)
    if split_totals is None:
        return False
    split_current, split_prior = split_totals
    comparisons = [
        parent == split
        for parent, split in (
            (parent_current, split_current),
            (parent_prior, split_prior),
        )
        if parent is not None and split is not None
    ]
    return bool(comparisons) and all(comparisons)


def _maturity_rows(table: ExtractedTable) -> list[tuple[int, str]]:
    return [
        (row_index, normalized_key(_cell(row, 0)))
        for row_index, row in enumerate(table.rows)
        if normalized_key(_cell(row, 0)) in {"ngan han", "dai han"}
    ]


def _maturity_split_totals(
    table: ExtractedTable,
) -> tuple[Decimal | None, Decimal | None] | None:
    maturity_rows = _maturity_rows(table)
    numeric_cols = [
        col
        for col in range(1, table.column_count)
        if any(_parse_number(_cell(row, col)) is not None for row in table.rows)
    ]
    if len(maturity_rows) != 2 or not numeric_cols:
        return None
    current_col = numeric_cols[0]
    prior_col = numeric_cols[-1] if len(numeric_cols) > 1 else numeric_cols[0]

    def period_total(column: int) -> Decimal | None:
        values = [
            _parse_number(_cell(table.rows[row_index], column))
            for row_index, _ in maturity_rows
        ]
        present = [value for value in values if value is not None]
        return sum(present, Decimal(0)) if present else None

    return period_total(current_col), period_total(prior_col)


def _reconcile_equity_movement_table(
    note_table: ExtractedTable,
    statement_items: list[StatementItem],
    tolerance: Decimal,
) -> list[NoteMatchResult] | None:
    """Đối chiếu bảng biến động vốn chủ sở hữu theo mẫu Thông tư 200."""
    recognition_text = normalized_key(
        " ".join(
            (note_table.title_hint, *note_table.context_hints)
            + tuple(_cell(row, 0) for row in note_table.rows[:3])
        )
    )
    if "von chu so huu" not in recognition_text:
        return None

    capital_col = _equity_column(note_table, "von gop cua chu so huu")
    retained_earnings_col = _equity_column(
        note_table,
        "loi nhuan sau thue chua phan phoi",
    )
    prior_closing_row = _equity_row(
        note_table,
        lambda label: (
            ("so cuoi nam truoc" in label and "dau nam nay" in label)
            or label in ("so cuoi nam truoc", "so dau nam nay", "so cuoi ky truoc")
        ),
    )
    closing_row = _equity_row(
        note_table,
        lambda label: any(kw in label for kw in ("so cuoi nam", "so cuoi ky", "so du cuoi nam", "so du cuoi ky")) and "truoc" not in label,
    )
    if (
        capital_col is None
        or retained_earnings_col is None
        or prior_closing_row is None
        or closing_row is None
    ):
        return None

    results = [
        _equity_note_match(
            note_table,
            statement_items,
            source_type=StatementType.BALANCE_SHEET_EQUITY,
            statement_code="411",
            fallback_name="Vốn góp của chủ sở hữu",
            current_row=closing_row,
            prior_row=prior_closing_row,
            note_col=capital_col,
            tolerance=tolerance,
            match_type="Tên dòng + mã BS",
        ),
        _equity_note_match(
            note_table,
            statement_items,
            source_type=StatementType.BALANCE_SHEET_EQUITY,
            statement_code="421",
            fallback_name="Lợi nhuận sau thuế chưa phân phối",
            current_row=closing_row,
            prior_row=prior_closing_row,
            note_col=retained_earnings_col,
            tolerance=tolerance,
            match_type="Tên dòng + mã BS",
        ),
    ]

    current_profit_row = _equity_row(
        note_table,
        lambda label: label == "loi nhuan thuan trong nam",
    )
    prior_profit_row = _equity_row(
        note_table,
        lambda label: "loi nhuan thuan trong nam truoc" in label,
    )
    if current_profit_row is not None or prior_profit_row is not None:
        results.append(
            _equity_note_match(
                note_table,
                statement_items,
                source_type=StatementType.INCOME_STATEMENT,
                statement_code="60",
                fallback_name="Lợi nhuận sau thuế thu nhập doanh nghiệp",
                current_row=current_profit_row,
                prior_row=prior_profit_row,
                note_col=retained_earnings_col,
                tolerance=tolerance,
                match_type="Tên dòng + PL",
            )
        )
    return results


def _equity_column(table: ExtractedTable, marker: str) -> int | None:
    header_limit = min(5, table.row_count)
    for col in range(1, table.column_count):
        header = normalized_key(" ".join(_cell(table.rows[row], col) for row in range(header_limit)))
        if marker in header:
            return col
    return None


def _equity_row(
    table: ExtractedTable,
    predicate: Callable[[str], bool],
) -> int | None:
    return next(
        (
            row_index
            for row_index, row in enumerate(table.rows)
            if predicate(normalized_key(_cell(row, 0)))
        ),
        None,
    )


def _equity_note_match(
    note_table: ExtractedTable,
    statement_items: list[StatementItem],
    *,
    source_type: StatementType,
    statement_code: str,
    fallback_name: str,
    current_row: int | None,
    prior_row: int | None,
    note_col: int,
    tolerance: Decimal,
    match_type: str,
) -> NoteMatchResult:
    note_current = (
        _parse_number(_cell(note_table.rows[current_row], note_col))
        if current_row is not None
        else None
    )
    note_prior = (
        _parse_number(_cell(note_table.rows[prior_row], note_col))
        if prior_row is not None
        else None
    )
    note_current_cell = (
        _excel_cell(current_row, note_col) if current_row is not None and note_current is not None else ""
    )
    note_prior_cell = (
        _excel_cell(prior_row, note_col) if prior_row is not None and note_prior is not None else ""
    )
    candidates = [
        item
        for item in statement_items
        if item.source_type == source_type and normalize_code(item.code) == statement_code
    ]
    item = _select_period_value_candidate(candidates, note_current, note_prior)
    if item is None:
        return NoteMatchResult(
            note_table_index=note_table.index,
            note_title=note_table.title_hint,
            source="PL" if source_type == StatementType.INCOME_STATEMENT else "BS",
            statement_code=statement_code,
            item_name=fallback_name,
            note_current=note_current,
            note_current_cell=note_current_cell,
            note_prior=note_prior,
            note_prior_cell=note_prior_cell,
            status="Statement not found",
            match_type=match_type,
            match_score=100.0,
        )

    current_difference = _difference(item.current_value, note_current)
    prior_difference = _difference(item.prior_value, note_prior)
    return NoteMatchResult(
        note_table_index=note_table.index,
        note_title=note_table.title_hint,
        statement_table_index=item.table_index,
        statement_current_cell=item.current_cell if item.current_value is not None else "",
        note_current_cell=note_current_cell,
        statement_prior_cell=item.prior_cell if item.prior_value is not None else "",
        note_prior_cell=note_prior_cell,
        source=item.source,
        statement_code=item.code,
        note_code=item.note_reference,
        item_name=item.name,
        statement_current=item.current_value,
        note_current=note_current,
        current_difference=current_difference,
        statement_prior=item.prior_value,
        note_prior=note_prior,
        prior_difference=prior_difference,
        status=_comparison_status(
            item.current_value,
            note_current,
            current_difference,
            item.prior_value,
            note_prior,
            prior_difference,
            tolerance,
        ),
        match_type=match_type,
        match_score=100.0,
    )


def _reconcile_contributed_capital_table(
    note_table: ExtractedTable,
    statement_items: list[StatementItem],
    tolerance: Decimal,
) -> list[NoteMatchResult] | None:
    context = normalized_key(" ".join((note_table.title_hint, *note_table.context_hints)))
    header = normalized_key(
        " ".join(value for row in note_table.rows[:3] for value in row)
    )
    if not any(
        marker in context
        for marker in ("von chu so huu", "von gop cua chu so huu")
    ) or "von da gop" not in header:
        return None

    current_cols, prior_cols = _note_period_columns(note_table)
    vnd_cols = {
        col
        for col in range(note_table.column_count)
        if "vnd" in normalized_key(
            " ".join(_cell(row, col) for row in note_table.rows[:3])
        )
    }
    current_vnd_cols, prior_vnd_cols = _capital_vnd_period_columns(
        current_cols,
        prior_cols,
        vnd_cols,
    )
    note_current, note_current_cell = _capital_period_total(note_table, current_vnd_cols)
    note_prior, note_prior_cell = _capital_period_total(note_table, prior_vnd_cols)
    candidates = [
        _with_inherited_note_reference(item, statement_items)
        for item in statement_items
        if item.source == "BS" and normalize_code(item.code) == "411"
    ]
    item = _select_period_value_candidate(candidates, note_current, note_prior)
    if item is None:
        return [
            NoteMatchResult(
                note_table_index=note_table.index,
                note_title=note_table.title_hint,
                source="BS",
                statement_code="411",
                item_name="Vốn góp của chủ sở hữu",
                note_current=note_current,
                note_current_cell=note_current_cell,
                note_prior=note_prior,
                note_prior_cell=note_prior_cell,
                status="Statement not found",
                match_type="Vốn đã góp + mã BS",
                match_score=100.0,
            )
        ]

    current_difference = _difference(item.current_value, note_current)
    prior_difference = _difference(item.prior_value, note_prior)
    return [
        NoteMatchResult(
            note_table_index=note_table.index,
            note_title=note_table.title_hint,
            statement_table_index=item.table_index,
            statement_current_cell=item.current_cell if item.current_value is not None else "",
            note_current_cell=note_current_cell,
            statement_prior_cell=item.prior_cell if item.prior_value is not None else "",
            note_prior_cell=note_prior_cell,
            source="BS",
            statement_code=item.code,
            note_code=item.note_reference,
            item_name=item.name,
            statement_current=item.current_value,
            note_current=note_current,
            current_difference=current_difference,
            statement_prior=item.prior_value,
            note_prior=note_prior,
            prior_difference=prior_difference,
            status=_comparison_status(
                item.current_value,
                note_current,
                current_difference,
                item.prior_value,
                note_prior,
                prior_difference,
                tolerance,
            ),
            match_type="Vốn đã góp + mã BS",
            match_score=100.0,
        )
    ]


def _capital_vnd_period_columns(
    current_cols: list[int],
    prior_cols: list[int],
    vnd_cols: set[int],
) -> tuple[list[int], list[int]]:
    current_anchor = min(current_cols) if current_cols else None
    prior_anchor = min(prior_cols) if prior_cols else None
    current_vnd_cols = [col for col in current_cols if col in vnd_cols]
    prior_vnd_cols = [col for col in prior_cols if col in vnd_cols]
    if not current_vnd_cols and current_anchor is not None:
        current_vnd_cols = [
            col
            for col in sorted(vnd_cols)
            if col >= current_anchor and (prior_anchor is None or col < prior_anchor)
        ]
    if not prior_vnd_cols and prior_anchor is not None:
        prior_vnd_cols = [col for col in sorted(vnd_cols) if col >= prior_anchor]
    return current_vnd_cols or current_cols, prior_vnd_cols or prior_cols


def _capital_period_total(
    table: ExtractedTable,
    columns: list[int],
) -> tuple[Decimal | None, str]:
    if not columns:
        return None, ""
    total_row = next(
        (
            row_index
            for row_index, row in enumerate(table.rows)
            if "tong cong" in normalized_key(_cell(row, 0))
        ),
        None,
    )
    if total_row is not None:
        return _note_row_period_value(table, total_row, columns)

    values: list[Decimal] = []
    first_cell = ""
    for row_index, row in enumerate(table.rows):
        if not clean_text(_cell(row, 0)):
            continue
        value, cell = _note_row_period_value(table, row_index, columns)
        if value is None:
            continue
        values.append(value)
        if not first_cell:
            first_cell = cell
    return (sum(values, Decimal(0)), first_cell) if values else (None, "")


def _reconcile_depreciation_expense_to_cash_flow(
    note_table: ExtractedTable,
    statement_items: list[StatementItem],
    tolerance: Decimal,
) -> list[NoteMatchResult] | None:
    normalized_title = _normalized_note_title(note_table)
    if not all(
        marker in normalized_title
        for marker in ("chi phi san xuat", "kinh doanh theo yeu to")
    ):
        return None
    row_index = next(
        (
            index
            for index, row in enumerate(note_table.rows)
            if normalized_key(_cell(row, 0))
            in {"chi phi khau hao tai san", "chi phi khau hao tai san co dinh"}
        ),
        None,
    )
    if row_index is None:
        return None

    current_cols, prior_cols = _note_period_columns(note_table)
    note_current, note_current_cell = _note_row_period_value(
        note_table,
        row_index,
        current_cols,
    )
    note_prior, note_prior_cell = _note_row_period_value(
        note_table,
        row_index,
        prior_cols,
    )
    candidates = [
        item
        for item in statement_items
        if item.source == "CF" and normalize_code(item.code) == "2"
    ]
    item = _select_absolute_period_value_candidate(candidates, note_current, note_prior)
    if item is None:
        return [
            NoteMatchResult(
                note_table_index=note_table.index,
                note_title=note_table.title_hint,
                source="CF",
                statement_code="02",
                item_name="Khấu hao tài sản cố định và bất động sản đầu tư",
                note_current=note_current,
                note_current_cell=note_current_cell,
                note_prior=note_prior,
                note_prior_cell=note_prior_cell,
                status="Statement not found",
                match_type="Tên dòng + mã CF",
                match_score=100.0,
            )
        ]

    current_difference = _absolute_difference(item.current_value, note_current)
    prior_difference = _absolute_difference(item.prior_value, note_prior)
    return [
        NoteMatchResult(
            note_table_index=note_table.index,
            note_title=note_table.title_hint,
            statement_table_index=item.table_index,
            statement_current_cell=item.current_cell if item.current_value is not None else "",
            note_current_cell=note_current_cell,
            statement_prior_cell=item.prior_cell if item.prior_value is not None else "",
            note_prior_cell=note_prior_cell,
            source="CF",
            statement_code=item.code,
            note_code=item.note_reference,
            item_name=item.name,
            statement_current=item.current_value,
            note_current=note_current,
            current_difference=current_difference,
            statement_prior=item.prior_value,
            note_prior=note_prior,
            prior_difference=prior_difference,
            status=_comparison_status(
                item.current_value,
                note_current,
                current_difference,
                item.prior_value,
                note_prior,
                prior_difference,
                tolerance,
            ),
            match_type="Tên dòng + mã CF",
            match_score=100.0,
        )
    ]


def _select_absolute_period_value_candidate(
    candidates: list[StatementItem],
    current: Decimal | None,
    prior: Decimal | None,
) -> StatementItem | None:
    if len(candidates) == 1:
        return candidates[0]
    exact = [
        item
        for item in candidates
        if all(
            abs(statement) == abs(note)
            for statement, note in (
                (item.current_value, current),
                (item.prior_value, prior),
            )
            if statement is not None and note is not None
        )
        and any(
            statement is not None and note is not None
            for statement, note in (
                (item.current_value, current),
                (item.prior_value, prior),
            )
        )
    ]
    return exact[0] if len(exact) == 1 else None


def _reconcile_state_tax_balance_rows(
    note_table: ExtractedTable,
    statement_items: list[StatementItem],
    tolerance: Decimal,
) -> list[NoteMatchResult] | None:
    row_mappings = {
        "thue va cac khoan khac phai thu nha nuoc": (
            "153",
            "Thuế và các khoản khác phải thu Nhà nước",
        ),
        "thue va cac khoan phai nop nha nuoc": (
            "313",
            "Thuế và các khoản phải nộp Nhà nước",
        ),
    }
    mapped_rows = [
        (row_index, *row_mappings[label])
        for row_index, row in enumerate(note_table.rows)
        if (label := normalized_key(_cell(row, 0))) in row_mappings
    ]
    if not mapped_rows:
        return None

    current_cols, prior_cols = _note_period_columns(note_table)
    balance_items = [item for item in statement_items if item.source == "BS"]
    results: list[NoteMatchResult] = []
    for row_index, code, fallback_name in mapped_rows:
        note_current, note_current_cell = _note_row_period_value(
            note_table,
            row_index,
            current_cols,
        )
        note_prior, note_prior_cell = _note_row_period_value(
            note_table,
            row_index,
            prior_cols,
        )
        candidates = [item for item in balance_items if normalize_code(item.code) == code]
        item = _select_period_value_candidate(candidates, note_current, note_prior)
        if item is None:
            results.append(
                NoteMatchResult(
                    note_table_index=note_table.index,
                    note_title=note_table.title_hint,
                    source="BS",
                    statement_code=code,
                    item_name=fallback_name,
                    note_current=note_current,
                    note_current_cell=note_current_cell,
                    note_prior=note_prior,
                    note_prior_cell=note_prior_cell,
                    status="Statement not found",
                    match_type="Tên dòng + mã BS",
                    match_score=100.0,
                )
            )
            continue

        current_difference = _difference(item.current_value, note_current)
        prior_difference = _difference(item.prior_value, note_prior)
        results.append(
            NoteMatchResult(
                note_table_index=note_table.index,
                note_title=note_table.title_hint,
                statement_table_index=item.table_index,
                statement_current_cell=item.current_cell if item.current_value is not None else "",
                note_current_cell=note_current_cell,
                statement_prior_cell=item.prior_cell if item.prior_value is not None else "",
                note_prior_cell=note_prior_cell,
                source="BS",
                statement_code=item.code,
                note_code=item.note_reference,
                item_name=item.name,
                statement_current=item.current_value,
                note_current=note_current,
                current_difference=current_difference,
                statement_prior=item.prior_value,
                note_prior=note_prior,
                prior_difference=prior_difference,
                status=_comparison_status(
                    item.current_value,
                    note_current,
                    current_difference,
                    item.prior_value,
                    note_prior,
                    prior_difference,
                    tolerance,
                ),
                match_type="Tên dòng + mã BS",
                match_score=100.0,
            )
        )
    return results


def _is_receivable_net_continuation_note(
    table: ExtractedTable,
    tables: list[ExtractedTable],
) -> bool:
    position = next(
        (index for index, candidate in enumerate(tables) if candidate.index == table.index),
        -1,
    )
    if position <= 0:
        return False
    previous = tables[position - 1]
    if (
        previous.statement_type != StatementType.NOTE
        or normalized_key(previous.title_hint) != normalized_key(table.title_hint)
    ):
        return False
    labels = [normalized_key(_cell(row, 0)) for row in table.rows]
    return (
        any(label.startswith("du phong phai thu") for label in labels)
        and any(label.startswith("gia tri thuan") for label in labels)
    )


def _reconcile_income_tax_schedule(
    note_table: ExtractedTable,
    statement_items: list[StatementItem],
    tolerance: Decimal,
) -> list[NoteMatchResult] | None:
    row_mappings = (
        (
            "loi nhuan truoc thue",
            ("loi nhuan ke toan truoc thue", "tong loi nhuan ke toan truoc thue"),
            "Tổng lợi nhuận kế toán trước thuế",
        ),
        (
            "chi phi thue thu nhap doanh nghiep hien hanh",
            ("chi phi thue thu nhap doanh nghiep hien hanh",),
            "Chi phí thuế thu nhập doanh nghiệp hiện hành",
        ),
    )
    mapped_rows = [
        (row_index, statement_markers, fallback_name)
        for row_index, row in enumerate(note_table.rows)
        for note_label, statement_markers, fallback_name in row_mappings
        if normalized_key(_cell(row, 0)) == note_label
    ]
    if not mapped_rows:
        return None

    current_cols, prior_cols = _note_period_columns(note_table)
    pl_items = [item for item in statement_items if item.source == "PL"]
    results: list[NoteMatchResult] = []
    for row_index, markers, fallback_name in mapped_rows:
        note_current, note_current_cell = _note_row_period_value(
            note_table,
            row_index,
            current_cols,
        )
        note_prior, note_prior_cell = _note_row_period_value(
            note_table,
            row_index,
            prior_cols,
        )
        candidates = [
            item
            for item in pl_items
            if any(marker in _normalized_item_name(item.name) for marker in markers)
        ]
        item = _select_period_value_candidate(candidates, note_current, note_prior)
        if item is None:
            results.append(
                NoteMatchResult(
                    note_table_index=note_table.index,
                    note_title=note_table.title_hint,
                    source="PL",
                    item_name=fallback_name,
                    note_current=note_current,
                    note_current_cell=note_current_cell,
                    note_prior=note_prior,
                    note_prior_cell=note_prior_cell,
                    status="Statement not found",
                    match_type="Tên dòng + PL",
                    match_score=100.0,
                )
            )
            continue
        current_difference = _difference(item.current_value, note_current)
        prior_difference = _difference(item.prior_value, note_prior)
        results.append(
            NoteMatchResult(
                note_table_index=note_table.index,
                note_title=note_table.title_hint,
                statement_table_index=item.table_index,
                statement_current_cell=item.current_cell if item.current_value is not None else "",
                note_current_cell=note_current_cell,
                statement_prior_cell=item.prior_cell if item.prior_value is not None else "",
                note_prior_cell=note_prior_cell,
                source="PL",
                statement_code=item.code,
                note_code=item.note_reference,
                item_name=item.name,
                statement_current=item.current_value,
                note_current=note_current,
                current_difference=current_difference,
                statement_prior=item.prior_value,
                note_prior=note_prior,
                prior_difference=prior_difference,
                status=_comparison_status(
                    item.current_value,
                    note_current,
                    current_difference,
                    item.prior_value,
                    note_prior,
                    prior_difference,
                    tolerance,
                ),
                match_type="Tên dòng + PL",
                match_score=100.0,
            )
        )
    return results


def _select_period_value_candidate(
    candidates: list[StatementItem],
    current: Decimal | None,
    prior: Decimal | None,
) -> StatementItem | None:
    if len(candidates) == 1:
        return candidates[0]
    exact = [
        item
        for item in candidates
        if all(
            statement == note
            for statement, note in (
                (item.current_value, current),
                (item.prior_value, prior),
            )
            if statement is not None and note is not None
        )
        and any(
            statement is not None and note is not None
            for statement, note in (
                (item.current_value, current),
                (item.prior_value, prior),
            )
        )
    ]
    return exact[0] if len(exact) == 1 else None


def _note_row_period_value(
    table: ExtractedTable,
    row_index: int,
    columns: list[int],
) -> tuple[Decimal | None, str]:
    for col in columns:
        value = _parse_number(_cell(table.rows[row_index], col))
        if value is not None:
            return value, _excel_cell(row_index, col)
    return None, ""


def _reconcile_maturity_rows(
    note_table: ExtractedTable,
    statement_items: list[StatementItem],
    tolerance: Decimal,
) -> list[NoteMatchResult] | None:
    maturity_rows = _maturity_rows(note_table)
    if len(maturity_rows) != 2:
        return None

    numeric_cols = [
        col
        for col in range(1, note_table.column_count)
        if any(_parse_number(_cell(row, col)) is not None for row in note_table.rows)
    ]
    if not numeric_cols:
        return None
    current_col = numeric_cols[0]
    prior_col = numeric_cols[-1] if len(numeric_cols) > 1 else numeric_cols[0]
    title = _maturity_base_title(_normalized_note_title(note_table))
    title_tokens = set(title.split())
    balance_items = [
        item
        for item in statement_items
        if item.source == "BS" and item.note_reference
    ]
    results: list[NoteMatchResult] = []

    for row_index, maturity in maturity_rows:
        required_tokens = title_tokens | set(maturity.split())
        candidates = [
            item
            for item in balance_items
            if required_tokens
            and required_tokens.issubset(set(_normalized_item_name(item.name).split()))
        ]
        item = candidates[0] if len(candidates) == 1 else None
        note_current = _parse_number(_cell(note_table.rows[row_index], current_col))
        note_prior = _parse_number(_cell(note_table.rows[row_index], prior_col))
        current_cell = _excel_cell(row_index, current_col) if note_current is not None else ""
        prior_cell = _excel_cell(row_index, prior_col) if note_prior is not None else ""
        if item is None:
            results.append(
                NoteMatchResult(
                    note_table_index=note_table.index,
                    note_title=note_table.title_hint,
                    source="BS",
                    item_name=f"{note_table.title_hint} - {_cell(note_table.rows[row_index], 0)}",
                    note_current=note_current,
                    note_current_cell=current_cell,
                    note_prior=note_prior,
                    note_prior_cell=prior_cell,
                    status="Statement not found",
                    match_type="Tên dòng và Thuyết minh",
                    match_score=100.0,
                )
            )
            continue

        current_difference = _difference(item.current_value, note_current)
        prior_difference = _difference(item.prior_value, note_prior)
        results.append(
            NoteMatchResult(
                note_table_index=note_table.index,
                note_title=note_table.title_hint,
                statement_table_index=item.table_index,
                statement_current_cell=item.current_cell if item.current_value is not None else "",
                note_current_cell=current_cell,
                statement_prior_cell=item.prior_cell if item.prior_value is not None else "",
                note_prior_cell=prior_cell,
                source="BS",
                statement_code=item.code,
                note_code=item.note_reference,
                item_name=item.name,
                statement_current=item.current_value,
                note_current=note_current,
                current_difference=current_difference,
                statement_prior=item.prior_value,
                note_prior=note_prior,
                prior_difference=prior_difference,
                status=_comparison_status(
                    item.current_value,
                    note_current,
                    current_difference,
                    item.prior_value,
                    note_prior,
                    prior_difference,
                    tolerance,
                ),
                match_type="Tên dòng và Thuyết minh",
                match_score=100.0,
            )
        )
    return results


def _maturity_base_title(title: str) -> str:
    combined_maturity_patterns = (
        r"\bngan\s*han\s*/\s*dai\s*han\b",
        r"\bdai\s*han\s*/\s*ngan\s*han\b",
    )
    normalized = title
    for pattern in combined_maturity_patterns:
        normalized = re.sub(pattern, " ", normalized)
    return " ".join(normalized.split())


def _normalized_note_title(table: ExtractedTable) -> str:
    hints = [table.title_hint, *table.context_hints]
    for hint in hints:
        normalized = re.sub(r"^\d+[.)]?\s*", "", normalized_key(hint))
        if normalized:
            return normalized
    return ""


def is_inventory_provision_rollforward_note(table: ExtractedTable) -> bool:
    """Bảng diễn biến tăng giảm dự phòng hàng tồn kho không đối chiếu trực tiếp với số dư BCTC."""
    text = normalized_key(
        " ".join([table.title_hint, *table.context_hints] + [value for row in table.rows for value in row])
    )
    is_provision = "du phong" in text or "du phong giam gia" in text
    has_rollforward = "trich lap" in text or "hoan nhap" in text or ("so dau nam" in text and "so cuoi nam" in text)
    return is_provision and has_rollforward


def _is_non_reconcilable_information_note(table: ExtractedTable) -> bool:
    if is_tax_rollforward_note(table) or is_inventory_provision_rollforward_note(table):
        return True
    text = normalized_key(
        " ".join(
            [table.title_hint, *table.context_hints]
            + [value for row in table.rows for value in row]
        )
    )
    if "so nam khau hao" in text or "thoi gian khau hao" in text:
        return True
    if _is_depreciation_life_range_table(table):
        return True
    if "khau hao het" in text and any(
        marker in text for marker in ("van con su dung", "dang con su dung")
    ):
        return True
    if _is_narrative_contract_information_table(table):
        return True
    if _is_contract_schedule_table(table):
        return True
    if "giao dich voi cac ben lien quan" in text or (
        "cac ben lien quan" in text and "duoc trinh bay o bang sau" in text
    ):
        return True
    if all(
        marker in text
        for marker in ("tien thue", "phai tra trong tuong lai", "hop dong thue")
    ):
        return True
    if any(
        marker in text
        for marker in ("the chap", "cam co", "dam bao cac khoan vay", "tai san dam bao")
    ) and _rows_with_numbers(table) <= 1:
        return True
    has_current_period = "so cuoi nam" in text or "nam nay" in text
    has_prior_period = "so dau nam" in text or "nam truoc" in text
    has_off_balance_marker = any(
        marker in text
        for marker in (
            "ngoai bang can doi",
            "ngoai te cac loai",
            "cac khoan muc ngoai bang",
        )
    )
    return (
        has_current_period
        and has_prior_period
        and has_off_balance_marker
        and _rows_with_numbers(table) == 1
    )


def is_tax_rollforward_note(table: ExtractedTable) -> bool:
    """Bảng diễn biến thuế không đối chiếu trực tiếp với một chỉ tiêu BCTC duy nhất."""
    header = normalized_key(
        " ".join(value for row in table.rows[:3] for value in row)
    )
    labels = [normalized_key(_cell(row, 0)) for row in table.rows]
    return (
        any(label.startswith("thue") for label in labels)
        and "so dau nam" in header
        and "so phai nop" in header
        and "so cuoi nam" in header
        and ("so da nop" in header or "khau tru" in header)
    )


def _is_depreciation_life_range_table(table: ExtractedTable) -> bool:
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
            marker in normalized_key(_cell(row, 0))
            for marker in (
                "nha cua",
                "may moc",
                "phuong tien van tai",
                "thiet bi van phong",
            )
        )
    )
    return has_year_header and range_count >= 2 and asset_class_count >= 2


def _is_narrative_contract_information_table(table: ExtractedTable) -> bool:
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
        "tong cong" in normalized_key(" ".join(row[:2])) for row in table.rows
    )
    return matches >= 3 and not has_explicit_total


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


def _rows_with_numbers(table: ExtractedTable) -> int:
    return sum(
        1
        for row in table.rows
        if any(_parse_number(value) is not None for value in row[1:])
    )


def _reconcile_asset_note_by_codes(
    note_table: ExtractedTable,
    statement_items: list[StatementItem],
    tolerance: Decimal,
) -> list[NoteMatchResult] | None:
    registry = _asset_note_registry(note_table)
    if registry is None:
        return None
    value_locations = _asset_note_value_locations(note_table)
    if value_locations is None:
        return None

    item_by_code = {
        normalize_code(item.code): item
        for item in statement_items
        if item.source_type == StatementType.BALANCE_SHEET_ASSETS
    }
    mappings = (
        (registry.cost_code, "nguyen_gia", "Nguyên giá"),
        (registry.depreciation_code, "hao_mon", "Giá trị hao mòn"),
        (registry.net_code, "con_lai", "Giá trị còn lại"),
    )
    results: list[NoteMatchResult] = []
    for code, group_key, fallback_name in mappings:
        item = item_by_code.get(code)
        current_row, current_col, prior_row, prior_col = value_locations[group_key]
        note_current = _parse_number(_cell(note_table.rows[current_row], current_col))
        note_prior = _parse_number(_cell(note_table.rows[prior_row], prior_col))
        note_current_cell = _excel_cell(current_row, current_col) if note_current is not None else ""
        note_prior_cell = _excel_cell(prior_row, prior_col) if note_prior is not None else ""
        if item is None:
            results.append(
                NoteMatchResult(
                    note_table_index=note_table.index,
                    note_title=note_table.title_hint,
                    source="BS",
                    statement_code=code,
                    item_name=fallback_name,
                    note_current=note_current,
                    note_current_cell=note_current_cell,
                    note_prior=note_prior,
                    note_prior_cell=note_prior_cell,
                    status="Statement not found",
                    match_type="Mã số BCTC",
                    match_score=100.0,
                )
            )
            continue
        current_difference = _difference(item.current_value, note_current)
        prior_difference = _difference(item.prior_value, note_prior)
        results.append(
            NoteMatchResult(
                note_table_index=note_table.index,
                note_title=note_table.title_hint,
                statement_table_index=item.table_index,
                statement_current_cell=item.current_cell if item.current_value is not None else "",
                note_current_cell=note_current_cell,
                statement_prior_cell=item.prior_cell if item.prior_value is not None else "",
                note_prior_cell=note_prior_cell,
                source=item.source,
                statement_code=item.code,
                note_code=item.note_reference,
                item_name=item.name,
                statement_current=item.current_value,
                note_current=note_current,
                current_difference=current_difference,
                statement_prior=item.prior_value,
                note_prior=note_prior,
                prior_difference=prior_difference,
                status=_comparison_status(
                    item.current_value,
                    note_current,
                    current_difference,
                    item.prior_value,
                    note_prior,
                    prior_difference,
                    tolerance,
                ),
                match_type="Mã số BCTC",
                match_score=100.0,
            )
        )
    return results


def _asset_note_value_locations(
    table: ExtractedTable,
) -> dict[str, tuple[int, int, int, int]] | None:
    row_groups = _asset_note_period_rows(table)
    total_col = _asset_note_total_column(table)
    if row_groups is not None and total_col is not None:
        return {
            key: (closing_row, total_col, opening_row, total_col)
            for key, (opening_row, closing_row) in row_groups.items()
        }
    return _compact_asset_note_value_locations(table)


def _compact_asset_note_value_locations(
    table: ExtractedTable,
) -> dict[str, tuple[int, int, int, int]] | None:
    current_col: int | None = None
    prior_col: int | None = None
    for col in range(1, table.column_count):
        header = normalized_key(" ".join(_cell(row, col) for row in table.rows[:3]))
        if current_col is None and ("so cuoi nam" in header or "nam nay" in header):
            current_col = col
        if prior_col is None and ("so dau nam" in header or "nam truoc" in header):
            prior_col = col
    if current_col is None or prior_col is None:
        return None

    rows: dict[str, int] = {}
    markers = (
        ("nguyen_gia", ("nguyen gia",)),
        (
            "hao_mon",
            ("khau hao luy ke", "hao mon luy ke", "gia tri khau hao", "gia tri hao mon"),
        ),
        ("con_lai", ("gia tri con lai",)),
    )
    for row_idx, row in enumerate(table.rows):
        label = normalized_key(_cell(row, 0))
        for key, prefixes in markers:
            if key not in rows and label.startswith(prefixes):
                rows[key] = row_idx
                break
    if set(rows) != {"nguyen_gia", "hao_mon", "con_lai"}:
        return None
    return {
        key: (row_idx, current_col, row_idx, prior_col)
        for key, row_idx in rows.items()
    }


def _asset_note_registry(table: ExtractedTable) -> AssetNoteCodeRegistry | None:
    hints = " ".join(_note_hints(table))
    return next(
        (
            registry
            for registry in ASSET_NOTE_CODE_REGISTRIES
            if any(marker in hints for marker in registry.title_markers)
        ),
        None,
    )


def _asset_note_total_column(table: ExtractedTable) -> int | None:
    for col in range(table.column_count - 1, 0, -1):
        header = normalized_key(" ".join(_cell(row, col) for row in table.rows[:4]))
        if "tong cong" in header:
            return col
    return None


def _asset_note_period_rows(
    table: ExtractedTable,
) -> dict[str, tuple[int, int]] | None:
    sections: list[tuple[str, int]] = []
    section_markers = (
        ("nguyen_gia", ("nguyen gia",)),
        ("hao_mon", ("gia tri hao mon", "gia tri khau hao")),
        ("con_lai", ("gia tri con lai",)),
    )
    for row_idx, row in enumerate(table.rows):
        label = normalized_key(_cell(row, 0))
        for key, markers in section_markers:
            if label.startswith(markers):
                sections.append((key, row_idx))
                break
    if [key for key, _ in sections] != ["nguyen_gia", "hao_mon", "con_lai"]:
        return None

    result: dict[str, tuple[int, int]] = {}
    for index, (key, section_row) in enumerate(sections):
        end_row = sections[index + 1][1] if index + 1 < len(sections) else len(table.rows)
        opening_row = _period_row(table, section_row + 1, end_row, opening=True)
        closing_row = _period_row(table, section_row + 1, end_row, opening=False)
        if opening_row is None or closing_row is None:
            return None
        result[key] = (opening_row, closing_row)
    return result


def _period_row(
    table: ExtractedTable,
    start_row: int,
    end_row: int,
    *,
    opening: bool,
) -> int | None:
    markers = ("so dau", "so du dau") if opening else ("so cuoi", "so du cuoi")
    for row_idx in range(start_row, end_row):
        label = normalized_key(_cell(table.rows[row_idx], 0))
        if label.startswith(markers):
            return row_idx
    return None


def normalize_code(value: object) -> str:
    text = clean_text(value)
    parsed = parse_accounting_number(text)
    if parsed is not None and parsed == parsed.to_integral_value():
        return str(int(parsed))
    return text.lower()


def _extract_statement_items(tables: list[ExtractedTable]) -> list[StatementItem]:
    items: list[StatementItem] = []
    for table in tables:
        if table.statement_type not in STATEMENT_TYPES:
            continue
        columns = _statement_columns(table)
        if columns is None:
            continue
        name_col, code_col, note_col, current_col, prior_col = columns
        for row_index, row in enumerate(table.rows):
            code = _cell(row, code_col)
            name = _cell(row, name_col)
            if not code or not name:
                continue
            items.append(
                StatementItem(
                    table_index=table.index,
                    source_type=table.statement_type,
                    code=code,
                    note_reference=_cell(row, note_col),
                    name=name,
                    current_value=_parse_number(_cell(row, current_col)),
                    prior_value=_parse_number(_cell(row, prior_col)),
                    current_cell=_excel_cell(row_index, current_col),
                    prior_cell=_excel_cell(row_index, prior_col),
                )
            )
    return items


def _statement_columns(table: ExtractedTable) -> tuple[int, int, int, int, int] | None:
    for row in table.rows[:8]:
        keys = [normalized_key(value) for value in row]
        code_col = _find_header(keys, "ma so")
        if code_col is None:
            continue
        note_col = _find_header(keys, "thuyet minh")
        numeric_cols = [
            col
            for col in range(code_col + 1, table.column_count)
            if col != note_col and any(_parse_number(_cell(data_row, col)) is not None for data_row in table.rows)
        ]
        if len(numeric_cols) < 2:
            return None
        name_col = next((col for col in range(code_col) if any(_cell(data_row, col) for data_row in table.rows)), 0)
        return name_col, code_col, note_col if note_col is not None else -1, numeric_cols[0], numeric_cols[1]
    return None


def _find_candidates(
    note_table: ExtractedTable,
    statement_items: list[StatementItem],
) -> list[tuple[StatementItem, str, float]]:
    hints = _note_hints(note_table)
    referenced_items = [item for item in statement_items if item.note_reference]
    semantic_candidate = _select_semantic_total_candidate(note_table, statement_items)
    if semantic_candidate is not None:
        return [
            (
                semantic_candidate,
                "Mã Thuyết minh + tên + Tổng cộng",
                100.0,
            )
        ]
    anchor_items = [item for item in referenced_items if _has_name_anchor(item.name, hints)]
    token_fallback = False
    if not anchor_items:
        anchor_items = [
            item
            for item in referenced_items
            if _has_token_name_anchor(item.name, hints)
        ]
        token_fallback = bool(anchor_items)

    if anchor_items:
        references = {normalized_key(item.note_reference) for item in anchor_items}
        candidates = [item for item in referenced_items if normalized_key(item.note_reference) in references]
        maturity = _title_maturity(note_table)
        if maturity:
            candidates = [
                item
                for item in candidates
                if maturity in _normalized_item_name(item.name)
            ]
        if token_fallback and not maturity:
            selected = _select_total_value_candidate(note_table, candidates)
            if selected is None:
                return []
            candidates = [selected]
        return [
            (
                item,
                (
                    "Tiêu đề kỳ hạn và Thuyết minh"
                    if maturity
                    else "Mã Thuyết minh + tên + Tổng cộng"
                    if token_fallback
                    else "Tham chiếu Thuyết minh"
                ),
                100.0,
            )
            for item in candidates
        ]
    return []


def _select_semantic_total_candidate(
    table: ExtractedTable,
    items: list[StatementItem],
) -> StatementItem | None:
    note_text = normalized_key(
        " ".join(
            [table.title_hint, *table.context_hints]
            + [value for row in table.rows[:4] for value in row]
        )
    )
    item_markers = next(
        (
            candidates
            for note_markers, candidates in NOTE_ITEM_SEMANTIC_MAPPINGS
            if all(marker in note_text for marker in note_markers)
        ),
        (),
    )
    if not item_markers:
        return None
    semantic_items = [
        _with_inherited_note_reference(item, items)
        for item in items
        if item.source == "BS"
        and any(marker in _normalized_item_name(item.name) for marker in item_markers)
    ]
    return _select_total_value_candidate(table, semantic_items)


def _with_inherited_note_reference(
    item: StatementItem,
    items: list[StatementItem],
) -> StatementItem:
    if item.note_reference:
        return item
    item_row = _statement_item_row(item)
    parent = max(
        (
            candidate
            for candidate in items
            if candidate.table_index == item.table_index
            and candidate.note_reference
            and _statement_item_row(candidate) < item_row
        ),
        key=_statement_item_row,
        default=None,
    )
    if parent is None:
        return item
    return replace(item, note_reference=parent.note_reference)


def _statement_item_row(item: StatementItem) -> int:
    match = re.search(r"\d+$", item.current_cell)
    return int(match.group()) if match else 0


def _has_token_name_anchor(item_name: str, hints: set[str]) -> bool:
    item_tokens = set(_normalized_item_name(item_name).split())
    return any(
        len(tokens) >= 3 and tokens.issubset(item_tokens)
        for hint in hints
        if (tokens := set(hint.split()))
    )


def _select_total_value_candidate(
    table: ExtractedTable,
    candidates: list[StatementItem],
) -> StatementItem | None:
    current, _, prior, _ = _note_total_vnd_values(table)
    comparable = [
        item
        for item in candidates
        if any(
            statement is not None and note is not None
            for statement, note in (
                (item.current_value, current),
                (item.prior_value, prior),
            )
        )
    ]
    exact = [
        item
        for item in comparable
        if all(
            statement == note
            for statement, note in (
                (item.current_value, current),
                (item.prior_value, prior),
            )
            if statement is not None and note is not None
        )
    ]
    if len(exact) == 1:
        return exact[0]
    return comparable[0] if len(comparable) == 1 else None


def _note_total_vnd_values(
    table: ExtractedTable,
) -> tuple[Decimal | None, str, Decimal | None, str]:
    total_row = next(
        (
            row_index
            for row_index, row in enumerate(table.rows)
            if "tong cong" in normalized_key(_cell(row, 0))
        ),
        None,
    )
    if total_row is None:
        return None, "", None, ""

    current_cols, prior_cols = _note_period_columns(table)
    vnd_cols = {
        col
        for col in range(1, table.column_count)
        if "vnd" in normalized_key(" ".join(_cell(row, col) for row in table.rows[:4]))
    }

    def value_for(columns: list[int]) -> tuple[Decimal | None, str]:
        eligible = [col for col in columns if not vnd_cols or col in vnd_cols]
        for col in reversed(eligible):
            value = _parse_number(_cell(table.rows[total_row], col))
            if value is not None:
                return value, _excel_cell(total_row, col)
        return None, ""

    current, current_cell = value_for(current_cols)
    prior, prior_cell = value_for(prior_cols)
    return current, current_cell, prior, prior_cell


def _title_maturity(table: ExtractedTable) -> str:
    title = re.sub(
        r"^\d+(?:\.\d+)*[.)]?\s*",
        "",
        normalized_key(table.title_hint),
    )
    return title if title in {"ngan han", "dai han"} else ""


def _expected_maturity_name(table: ExtractedTable, maturity: str) -> str:
    parent = next(
        (
            clean_text(hint)
            for hint in table.context_hints
            if maturity not in normalized_key(hint)
        ),
        "",
    )
    parent = re.sub(r"^\d+(?:\.\d+)*[.)]?\s*", "", parent).strip()
    label = "ngắn hạn" if maturity == "ngan han" else "dài hạn"
    return f"{parent} {label}".strip()


def _note_total_values(
    table: ExtractedTable,
) -> tuple[Decimal | None, str, Decimal | None, str]:
    current_cols, prior_cols = _note_period_columns(table)
    total_row = next(
        (
            row_index
            for row_index, row in enumerate(table.rows)
            if "tong cong" in normalized_key(_cell(row, 0))
        ),
        None,
    )
    if total_row is None:
        return None, "", None, ""

    def period_value(columns: list[int]) -> tuple[Decimal | None, str]:
        for col in columns:
            value = _parse_number(_cell(table.rows[total_row], col))
            if value is not None:
                return value, _excel_cell(total_row, col)
        return None, ""

    current_value, current_cell = period_value(current_cols)
    prior_value, prior_cell = period_value(prior_cols)
    return current_value, current_cell, prior_value, prior_cell


def _compare_candidate(
    note_table: ExtractedTable,
    item: StatementItem,
    match_type: str,
    score: float,
    tolerance: Decimal,
) -> NoteMatchResult:
    current_col, prior_col = _note_period_columns(note_table)
    direct_name_match = _normalized_item_name(item.name) == normalized_key(note_table.title_hint)
    can_report_difference = _has_reconciliation_total(note_table) and (
        direct_name_match
        or match_type
        in {
            "Tiêu đề kỳ hạn và Thuyết minh",
            "Mã Thuyết minh + tên + Tổng cộng",
        }
    )
    note_current, note_current_cell = _closest_note_value(
        note_table,
        current_col,
        item.current_value,
        tolerance,
        allow_difference=can_report_difference,
    )
    note_prior, note_prior_cell = _closest_note_value(
        note_table,
        prior_col,
        item.prior_value,
        tolerance,
        allow_difference=can_report_difference,
    )
    current_difference = _difference(item.current_value, note_current)
    prior_difference = _difference(item.prior_value, note_prior)
    status = _comparison_status(
        item.current_value,
        note_current,
        current_difference,
        item.prior_value,
        note_prior,
        prior_difference,
        tolerance,
    )
    return NoteMatchResult(
        note_table_index=note_table.index,
        note_title=note_table.title_hint,
        statement_table_index=item.table_index,
        statement_current_cell=item.current_cell if item.current_value is not None else "",
        note_current_cell=note_current_cell,
        statement_prior_cell=item.prior_cell if item.prior_value is not None else "",
        note_prior_cell=note_prior_cell,
        source=item.source,
        statement_code=item.code,
        note_code=item.note_reference,
        item_name=item.name,
        statement_current=item.current_value,
        note_current=note_current,
        current_difference=current_difference,
        statement_prior=item.prior_value,
        note_prior=note_prior,
        prior_difference=prior_difference,
        status=status,
        match_type=match_type,
        match_score=round(score, 1),
    )


def _note_hints(table: ExtractedTable) -> set[str]:
    values = [table.title_hint, *table.context_hints]
    return {
        normalized
        for value in values
        if (
            normalized := re.sub(
                r"^\d+(?:\.\d+)*[.)]?\s*",
                "",
                normalized_key(value),
            )
        )
    }


def _note_period_columns(table: ExtractedTable) -> tuple[list[int], list[int]]:
    current_cols, prior_cols = detect_period_columns_from_headers(table)
    if current_cols or prior_cols:
        return current_cols, prior_cols
    numeric_cols = [
        col
        for col in range(table.column_count)
        if any(_parse_number(_cell(row, col)) is not None for row in table.rows)
    ]
    if len(numeric_cols) >= 2:
        return [numeric_cols[0]], [numeric_cols[-1]]
    return numeric_cols, numeric_cols


def _closest_note_value(
    table: ExtractedTable,
    columns: list[int],
    target: Decimal | None,
    tolerance: Decimal,
    *,
    allow_difference: bool,
) -> tuple[Decimal | None, str]:
    if target is None or not columns:
        return None, ""
    values = [
        (value, _excel_cell(row_index, col))
        for row_index, row in enumerate(table.rows)
        for col in columns
        if (value := _parse_number(_cell(row, col))) is not None
    ]
    if not values:
        return None, ""
    closest, cell = min(values, key=lambda candidate: abs(target - candidate[0]))
    if allow_difference or abs(target - closest) <= tolerance:
        return closest, cell
    return None, ""


def _excel_cell(row_index: int, col_index: int) -> str:
    """Chuyển tọa độ 0-based trong bảng trích xuất sang ô ở sheet chi tiết."""
    column_number = col_index + 1
    letters = ""
    while column_number:
        column_number, remainder = divmod(column_number - 1, 26)
        letters = chr(65 + remainder) + letters
    return f"{letters}{row_index + DEFAULT_EXCEL_DATA_START_ROW}"


def _comparison_status(
    statement_current: Decimal | None,
    note_current: Decimal | None,
    current_difference: Decimal | None,
    statement_prior: Decimal | None,
    note_prior: Decimal | None,
    prior_difference: Decimal | None,
    tolerance: Decimal,
) -> str:
    comparisons = [
        difference
        for statement, note, difference in (
            (statement_current, note_current, current_difference),
            (statement_prior, note_prior, prior_difference),
        )
        if statement is not None and note is not None and difference is not None
    ]
    if not comparisons:
        return "Note value not found"
    return "Matched" if all(abs(difference) <= tolerance for difference in comparisons) else "Difference"


def _difference(statement_value: Decimal | None, note_value: Decimal | None) -> Decimal | None:
    if statement_value is None or note_value is None:
        return None
    return (statement_value - note_value).quantize(
        RECONCILIATION_ROUND_QUANTUM,
        rounding=ROUND_HALF_UP,
    )


def _absolute_difference(
    statement_value: Decimal | None,
    note_value: Decimal | None,
) -> Decimal | None:
    if statement_value is None or note_value is None:
        return None
    return (abs(statement_value) - abs(note_value)).quantize(
        RECONCILIATION_ROUND_QUANTUM,
        rounding=ROUND_HALF_UP,
    )


def _normalized_item_name(value: str) -> str:
    return re.sub(r"^\d+[.)]?\s*", "", normalized_key(value))


def _has_name_anchor(item_name: str, hints: set[str]) -> bool:
    normalized_name = _normalized_item_name(item_name)
    return any(
        hint == normalized_name
        or (len(hint) >= 10 and hint in normalized_name)
        or (len(normalized_name) >= 10 and normalized_name in hint)
        for hint in hints
    )


def _has_reconciliation_total(table: ExtractedTable) -> bool:
    title = normalized_key(table.title_hint)
    for row in table.rows:
        label = normalized_key(_cell(row, 0))
        if not label:
            continue
        if any(marker in label for marker in ("tong cong", "gia tri thuan", "so cuoi nam")):
            return True
        if len(label) >= 10 and label in title:
            return True
    return False


def _parse_number(value: object) -> Decimal | None:
    text = clean_text(value)
    if text in DASH_ZERO_MARKERS:
        return Decimal(0)
    return parse_accounting_number(text)


def _find_header(keys: list[str], expected: str) -> int | None:
    return next((index for index, key in enumerate(keys) if key == expected or expected in key), None)


def _cell(row: list[str], index: int) -> str:
    if index < 0 or index >= len(row):
        return ""
    return clean_text(row[index])
