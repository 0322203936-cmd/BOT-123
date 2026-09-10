from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook

from inventory_box_transform import DATE_NUMBER_FORMAT, normalize_date_formats


class InventoryBoxesTests(unittest.TestCase):
    def test_normalizes_dates_only_on_the_availability_sheet(self) -> None:
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.xlsx"
            output = Path(temp_dir) / "normalized.xlsx"
            workbook = Workbook()
            active = workbook.active
            active.title = "Customer View"
            active["A1"] = "Debe conservarse"
            availability = workbook.create_sheet("Availability")
            availability.append(["Product Description", "Available From"])
            availability.append(["A", date(2026, 9, 14)])
            availability["B2"].number_format = "m/d/yy"
            workbook.save(source)
            workbook.close()

            normalize_date_formats(source, output)

            result = load_workbook(output, data_only=False)
            try:
                self.assertEqual(result["Availability"]["B2"].number_format, DATE_NUMBER_FORMAT)
                self.assertEqual(result["Customer View"]["A1"].value, "Debe conservarse")
            finally:
                result.close()


if __name__ == "__main__":
    unittest.main()
