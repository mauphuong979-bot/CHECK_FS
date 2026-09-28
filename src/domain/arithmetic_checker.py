from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP

from src.domain.audit_rules import ArithmeticRule
from src.domain.models import TableValue, TableValueStatus
from src.extraction.table_values import normalize_statement_code


CALCULABLE_VALUE_STATUSES = {
    TableValueStatus.NUMBER,
    TableValueStatus.ZERO,
    TableValueStatus.DASH_ZERO,
}
CHECK_ROUND_QUANTUM = Decimal("0.01")


@dataclass(frozen=True)
class ArithmeticCheckValue:
    difference: Decimal | None
    review_codes: tuple[str, ...] = ()
    applicable: bool = True

    @property
    def needs_review(self) -> bool:
        return bool(self.review_codes)


def calculate_same_table_arithmetic(
    rule: ArithmeticRule,
    value_column: int,
    values: dict[str, dict[int, TableValue]],
) -> ArithmeticCheckValue:
    review_codes: list[str] = []
    target = _required_calculation_value(
        rule.target_code,
        value_column,
        values,
        review_codes,
    )
    total = Decimal(0)
    present_term_count = 0

    for term in rule.terms:
        if term.source is not None:
            raise ValueError("Rule đối chiếu chéo không thuộc phép tính nội bảng.")
        normalized_code = normalize_statement_code(term.code)
        if normalized_code not in values:
            continue
        present_term_count += 1
        value = _required_calculation_value(
            term.code,
            value_column,
            values,
            review_codes,
        )
        if value is not None:
            total += Decimal(term.coefficient) * value

    if present_term_count == 0:
        return ArithmeticCheckValue(None, applicable=False)
    if review_codes or target is None:
        return ArithmeticCheckValue(None, tuple(dict.fromkeys(review_codes)))
    difference = (target - total).quantize(
        CHECK_ROUND_QUANTUM,
        rounding=ROUND_HALF_UP,
    )
    return ArithmeticCheckValue(difference)


def _required_calculation_value(
    code: str,
    value_column: int,
    values: dict[str, dict[int, TableValue]],
    review_codes: list[str],
) -> Decimal | None:
    normalized_code = normalize_statement_code(code)
    cell_value = values.get(normalized_code, {}).get(value_column)
    if (
        cell_value is None
        or cell_value.status not in CALCULABLE_VALUE_STATUSES
        or cell_value.value is None
    ):
        review_codes.append(code)
        return None
    return cell_value.value
