import unittest

from reunion_range_shift import (
    column_letter,
    find_cor_columns,
    find_last_data_row,
    next_excel_date,
    shifted_column_payload,
    cleared_last_column_payload,
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

    def test_shift_translates_subtotal_formulas(self) -> None:
        values = [[100], [40], [60]]
        formulas = [["=+K4+K5"], [40], [60]]
        self.assertEqual(
            shifted_column_payload(values, formulas, 11, 5),
            [["=+E4+E5"], [40], [60]],
        )

    def test_last_column_keeps_subtotals_and_clears_inputs(self) -> None:
        formulas = [["=+BG4+BG5"], [40], [60]]
        self.assertEqual(
            cleared_last_column_payload(formulas, 3),
            [["=+BG4+BG5"], [0], [0]],
        )


if __name__ == "__main__":
    unittest.main()
