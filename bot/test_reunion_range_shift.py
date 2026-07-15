import unittest

from reunion_range_shift import (
    column_letter,
    find_cor_columns,
    find_last_data_row,
    next_excel_date,
)


class ReunionRangeShiftTests(unittest.TestCase):
    def test_column_letters(self) -> None:
        self.assertEqual(column_letter(5), "E")
        self.assertEqual(column_letter(53), "BA")
        self.assertEqual(column_letter(59), "BG")

    def test_detects_all_cor_columns(self) -> None:
        headers = [""] * 59
        for column in (5, 11, 17, 23, 29, 35, 41, 47, 53, 59):
            headers[column - 1] = "Cor"
        self.assertEqual(
            find_cor_columns(headers),
            [5, 11, 17, 23, 29, 35, 41, 47, 53, 59],
        )

    def test_rejects_irregular_cor_blocks(self) -> None:
        with self.assertRaises(RuntimeError):
            find_cor_columns(["", "Cor", "", "Cor"])

    def test_data_rows_ignore_calculations_below(self) -> None:
        values = [["Flor A"], ["Flor B"], [""], [""]]
        self.assertEqual(find_last_data_row(values, start_row=3), 4)

    def test_date_advances_one_day(self) -> None:
        self.assertEqual(next_excel_date(46218), 46219)
        self.assertEqual(next_excel_date("7/15/2026"), 46219)


if __name__ == "__main__":
    unittest.main()
