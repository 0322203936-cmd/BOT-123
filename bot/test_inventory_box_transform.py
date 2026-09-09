from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

try:
    from inventory_box_transform import (
        apply_inventory_rules,
        transform_inventory_workbook,
    )
except ImportError:
    try:
        from bot.inventory_box_transform import (
            apply_inventory_rules,
            transform_inventory_workbook,
        )
    except ImportError:
        apply_inventory_rules = None
        transform_inventory_workbook = None


class InventoryBoxTransformTests(unittest.TestCase):
    assumed_today = date(2026, 9, 10)

    def test_removes_window_and_sundays_and_adds_one_row_per_product(self) -> None:
        rows = [
            ["Pacific", "A", "Bunch", date(2026, 9, 9)],
            ["Pacific", "A", "Bunch", date(2026, 9, 10)],
            ["Pacific", "A", "Bunch", date(2026, 9, 14)],
            ["Pacific", "B", "Bunch", date(2026, 9, 13)],
            ["Pacific", "B", "Bunch", date(2026, 9, 15)],
        ]

        self.assertIsNotNone(apply_inventory_rules)
        result = apply_inventory_rules(
            rows,
            product_column=1,
            date_column=3,
            assumed_today=self.assumed_today,
        )

        output_dates = [row.values[3] for row in result.output_rows]
        self.assertEqual(
            output_dates,
            [
                date(2026, 9, 9),
                date(2026, 9, 14),
                date(2026, 9, 15),
                date(2026, 9, 15),
                date(2026, 9, 16),
            ],
        )
        added = [row for row in result.output_rows if row.is_added]
        self.assertEqual(len(added), 2)
        self.assertEqual([row.values[1] for row in added], ["A", "B"])
        self.assertEqual(result.removed_rows, 2)
        self.assertEqual(result.final_sunday_rows, 0)

    def test_saturday_latest_row_adds_monday(self) -> None:
        rows = [["Pacific", "A", "Bunch", date(2026, 9, 19)]]

        self.assertIsNotNone(apply_inventory_rules)
        result = apply_inventory_rules(
            rows,
            product_column=1,
            date_column=3,
            assumed_today=self.assumed_today,
        )

        self.assertEqual(result.output_rows[-1].values[3], date(2026, 9, 21))

    def test_rejects_a_product_with_no_surviving_dated_row(self) -> None:
        rows = [["Pacific", "A", "Bunch", date(2026, 9, 10)]]

        self.assertIsNotNone(apply_inventory_rules)
        with self.assertRaisesRegex(RuntimeError, "A"):
            apply_inventory_rules(
                rows,
                product_column=1,
                date_column=3,
                assumed_today=self.assumed_today,
            )

    def test_workbook_transform_copies_format_to_added_row(self) -> None:
        self.assertIsNotNone(transform_inventory_workbook)
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.xlsx"
            output = Path(temp_dir) / "output.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Availability"
            sheet.append(["Vendor Name", "Product Description", "Package Type", "Available From"])
            sheet.append(["Pacific", "A", "Bunch", date(2026, 9, 14)])
            sheet["B2"].fill = PatternFill(fill_type="solid", fgColor="FFFF00")
            sheet["B2"].font = Font(name="Arial", bold=True)
            sheet["D2"].number_format = r"yyyy\-mm\-dd"
            workbook.save(source)
            workbook.close()

            transform_inventory_workbook(
                source,
                output,
                assumed_today=self.assumed_today,
            )

            result = load_workbook(output, data_only=False)
            try:
                sheet = result.active
                self.assertEqual(sheet.max_row, 3)
                self.assertEqual(sheet["B3"].value, "A")
                self.assertEqual(sheet["D3"].value.date(), date(2026, 9, 15))
                self.assertEqual(sheet["B3"]._style, sheet["B2"]._style)
                self.assertEqual(sheet["D3"].number_format, sheet["D2"].number_format)
            finally:
                result.close()


if __name__ == "__main__":
    unittest.main()
