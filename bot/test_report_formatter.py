import unittest

from report_formatter import corrected_veronica_txr


class CorrectedVeronicaTxrTests(unittest.TestCase):
    def test_bouquet_is_always_eight(self) -> None:
        self.assertEqual(corrected_veronica_txr("BOUQUET", -1.0), 8)
        self.assertEqual(corrected_veronica_txr(" bouquet ", 10.0), 8)

    def test_other_categories_are_ten(self) -> None:
        self.assertEqual(corrected_veronica_txr("CB / BULK", 10.0), 10)
        self.assertEqual(corrected_veronica_txr("CB / BULK", -1.0), 10)
        self.assertEqual(corrected_veronica_txr(None, None), 10)


if __name__ == "__main__":
    unittest.main()
