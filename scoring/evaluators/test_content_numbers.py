import unittest

try:
    from .content import _numbers_in_text
except ImportError:  # pragma: no cover
    from content import _numbers_in_text


class NumbersInTextTests(unittest.TestCase):
    def test_french_price_with_space_grouping(self):
        value = "Une configuration performante proposée au prix de 4 999,00 €."
        self.assertEqual(_numbers_in_text(value), [4999.0])

    def test_non_breaking_space_grouping(self):
        self.assertEqual(_numbers_in_text("4\u202f999,00 €"), [4999.0])

    def test_us_and_european_mixed_separators(self):
        self.assertEqual(_numbers_in_text("$4,999.00 / 4.999,00 €"), [4999.0, 4999.0])

    def test_decimal_measurements_remain_separate_numbers(self):
        self.assertEqual(_numbers_in_text("183 x 61 x 1,5 cm"), [183.0, 61.0, 1.5])

    def test_level_prefix_is_not_a_decimal(self):
        self.assertEqual(_numbers_in_text("reaches Lv.50."), [50.0])
        self.assertEqual(_numbers_in_text("costs .5 gold"), [0.5])


if __name__ == "__main__":
    unittest.main()
