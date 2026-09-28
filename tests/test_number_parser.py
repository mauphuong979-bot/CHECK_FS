from decimal import Decimal
import unittest

from src.normalization.number_parser import parse_accounting_number


def test_parse_vietnamese_thousand_separator():
    assert parse_accounting_number("1.628.000.000") == Decimal("1628000000")


def test_parse_accounting_negative():
    assert parse_accounting_number("(1.000)") == Decimal("-1000")


def test_ignore_non_number_text():
    assert parse_accounting_number("Tổng tài sản") is None


class NumberParserUnittestTest(unittest.TestCase):
    def test_existing_parser_cases(self):
        test_parse_vietnamese_thousand_separator()
        test_parse_accounting_negative()
        test_ignore_non_number_text()
