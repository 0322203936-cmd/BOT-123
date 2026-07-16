import unittest

from report_formatter import corrected_veronica_txr


class CorrectedVeronicaTxrTests(unittest.TestCase):
    def test_preserves_existing_eight_and_ten(self) -> None:
        self.assertEqual(corrected_veronica_txr("VERONICA", 8.0), 8)
        self.assertEqual(corrected_veronica_txr("VERONICA", 10.0), 10)

    def test_converts_other_values_to_ten(self) -> None:
        self.assertEqual(corrected_veronica_txr("VERONICA", -1.0), 10)
        self.assertEqual(corrected_veronica_txr("VERONICA", None), 10)

    def test_keeps_wildflower_pk_8_exception_at_eight(self) -> None:
        self.assertEqual(
            corrected_veronica_txr("NCP WILDFLOWER BOUQUET PK 8", -1.0),
            8,
        )


if __name__ == "__main__":
    unittest.main()
