from decimal import Decimal
import unittest

from src.domain.arithmetic_checker import calculate_same_table_arithmetic
from src.domain.audit_rules import ArithmeticRule, RuleTerm
from src.domain.models import StatementType, TableValue, TableValueStatus


class ArithmeticCheckerTest(unittest.TestCase):
    def setUp(self):
        self.rule = ArithmeticRule("100", (RuleTerm("110"), RuleTerm("120")))

    def test_calculates_with_decimal_and_explicit_zero_states(self):
        values = {
            "100": {3: TableValue(TableValueStatus.NUMBER, Decimal("100.25"))},
            "110": {3: TableValue(TableValueStatus.NUMBER, Decimal("100.25"))},
            "120": {3: TableValue(TableValueStatus.ZERO, Decimal("0"))},
        }

        result = calculate_same_table_arithmetic(self.rule, 3, values)

        self.assertEqual(result.difference, Decimal("0.00"))
        self.assertFalse(result.needs_review)

    def test_dash_zero_is_calculable(self):
        values = {
            "100": {3: TableValue(TableValueStatus.NUMBER, Decimal("100"))},
            "110": {3: TableValue(TableValueStatus.NUMBER, Decimal("100"))},
            "120": {3: TableValue(TableValueStatus.DASH_ZERO, Decimal("0"))},
        }

        result = calculate_same_table_arithmetic(self.rule, 3, values)

        self.assertEqual(result.difference, Decimal("0"))
        self.assertFalse(result.needs_review)

    def test_missing_and_invalid_values_require_review(self):
        values = {
            "100": {3: TableValue(TableValueStatus.NUMBER, Decimal("100"))},
            "110": {3: TableValue(TableValueStatus.MISSING)},
            "120": {3: TableValue(TableValueStatus.INVALID)},
        }

        result = calculate_same_table_arithmetic(self.rule, 3, values)

        self.assertIsNone(result.difference)
        self.assertEqual(result.review_codes, ("110", "120"))

    def test_absent_code_requires_review(self):
        values = {
            "110": {3: TableValue(TableValueStatus.NUMBER, Decimal("100"))},
        }
        result = calculate_same_table_arithmetic(self.rule, 3, values)

        self.assertIsNone(result.difference)
        self.assertEqual(result.review_codes, ("100",))

    def test_absent_component_code_is_treated_as_not_applicable_term(self):
        values = {
            "100": {3: TableValue(TableValueStatus.NUMBER, Decimal("100"))},
            "110": {3: TableValue(TableValueStatus.NUMBER, Decimal("100"))},
        }

        result = calculate_same_table_arithmetic(self.rule, 3, values)

        self.assertEqual(result.difference, Decimal("0"))
        self.assertFalse(result.needs_review)

    def test_rule_is_not_applicable_when_no_component_code_exists(self):
        values = {
            "100": {3: TableValue(TableValueStatus.NUMBER, Decimal("100"))},
        }

        result = calculate_same_table_arithmetic(self.rule, 3, values)

        self.assertFalse(result.applicable)
        self.assertFalse(result.needs_review)

    def test_rejects_cross_statement_rule(self):
        rule = ArithmeticRule(
            "100",
            (RuleTerm("110", source=StatementType.INCOME_STATEMENT),),
        )

        with self.assertRaisesRegex(ValueError, "đối chiếu chéo"):
            calculate_same_table_arithmetic(rule, 3, {})

    def test_differences_are_exact_without_rounding_tolerance(self):
        one_cent = {
            "100": {3: TableValue(TableValueStatus.NUMBER, Decimal("100.01"))},
            "110": {3: TableValue(TableValueStatus.NUMBER, Decimal("100"))},
        }
        sub_cent = {
            "100": {3: TableValue(TableValueStatus.NUMBER, Decimal("100.004"))},
            "110": {3: TableValue(TableValueStatus.NUMBER, Decimal("100"))},
        }
        exact = {
            "100": {3: TableValue(TableValueStatus.NUMBER, Decimal("100.10"))},
            "110": {3: TableValue(TableValueStatus.NUMBER, Decimal("100.1"))},
        }

        self.assertEqual(
            calculate_same_table_arithmetic(self.rule, 3, one_cent).difference,
            Decimal("0.01"),
        )
        self.assertEqual(
            calculate_same_table_arithmetic(self.rule, 3, sub_cent).difference,
            Decimal("0.004"),
        )
        self.assertEqual(
            calculate_same_table_arithmetic(self.rule, 3, exact).difference,
            Decimal(0),
        )

if __name__ == "__main__":
    unittest.main()
