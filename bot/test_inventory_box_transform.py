from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table
from openpyxl.styles import Font, PatternFill

try:
    from inventory_box_transform import (
        DATE_NUMBER_FORMAT,
        apply_inventory_rules,
        create_inventory_email_workbook,
        create_single_sheet_workbook,
        rebuild_customer_view_from_availability,
        transform_inventory_workbook,
        refresh_workbook_with_komet_inventory,
)
except ImportError:
    try:
        from bot.inventory_box_transform import (
            DATE_NUMBER_FORMAT,
            apply_inventory_rules,
            create_inventory_email_workbook,
            create_single_sheet_workbook,
            rebuild_customer_view_from_availability,
            transform_inventory_workbook,
            refresh_workbook_with_komet_inventory,
        )
    except ImportError:
        apply_inventory_rules = None
        create_single_sheet_workbook = None
        DATE_NUMBER_FORMAT = None
        create_inventory_email_workbook = None
        rebuild_customer_view_from_availability = None
        transform_inventory_workbook = None
        refresh_workbook_with_komet_inventory = None
try:
    from inventory_box_transform import _inventory_rows_from_xls
except ImportError:
    _inventory_rows_from_xls = None


class InventoryBoxTransformTests(unittest.TestCase):
    @patch("pandas.read_excel")
    def test_reads_legacy_xls_inventory_export(self, read_excel):
        read_excel.return_value = pd.DataFrame([
            ["Pricing", None, None],
            ["AWB", "Product", "Aging", "Qty"],
            ["000-2026-0911", "Rose Red", -2, 4],
        ])

        rows, header_row, headers = _inventory_rows_from_xls(Path("pricing.xls"))

        self.assertEqual(header_row, 2)
        self.assertEqual(headers["product"], 2)
        self.assertEqual(rows[0]["product"], "Rose Red")
        self.assertEqual(rows[0]["qty"], 4)

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
                date(2026, 9, 14),
                date(2026, 9, 15),
                date(2026, 9, 15),
                date(2026, 9, 16),
            ],
        )
        added = [row for row in result.output_rows if row.is_added]
        self.assertEqual(len(added), 2)
        self.assertEqual([row.values[1] for row in added], ["A", "B"])
        self.assertEqual(result.removed_rows, 3)
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

    def test_creates_a_komet_copy_with_only_the_availability_sheet(self) -> None:
        self.assertIsNotNone(create_single_sheet_workbook)
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.xlsx"
            output = Path(temp_dir) / "komet.xlsx"
            workbook = Workbook()
            active = workbook.active
            active.title = "Customer View"
            active["A1"] = "No debe ir a Komet"
            availability = workbook.create_sheet("Availability")
            availability.append(["Product Description", "Available From"])
            availability.append(["A", date(2026, 9, 14)])
            availability["A2"].font = Font(name="Arial", bold=True)
            availability["B2"].number_format = "mmm d, yyyy"
            workbook.save(source)
            workbook.close()

            create_single_sheet_workbook(
                source,
                output,
                normalize_available_from_dates=True,
            )

            result = load_workbook(output, data_only=False)
            try:
                self.assertEqual(result.sheetnames, ["Availability"])
                self.assertEqual(result.active["A2"].value, "A")
                self.assertEqual(result.active["A2"]._style, availability["A2"]._style)
                self.assertEqual(result.active["B2"].value.date(), date(2026, 9, 14))
                self.assertEqual(result.active["B2"].number_format, DATE_NUMBER_FORMAT)
            finally:
                result.close()

    def test_default_single_sheet_copy_does_not_change_date_format(self) -> None:
        self.assertIsNotNone(create_single_sheet_workbook)
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.xlsx"
            output = Path(temp_dir) / "copy.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Availability"
            sheet.append(["Product Description", "Available From"])
            sheet.append(["A", date(2026, 9, 14)])
            sheet["B2"].number_format = "mmm d, yyyy"
            workbook.save(source)
            workbook.close()

            create_single_sheet_workbook(source, output)

            result = load_workbook(output, data_only=False)
            try:
                self.assertEqual(result["Availability"]["B2"].number_format, "mmm d, yyyy")
            finally:
                result.close()

    def test_creates_an_email_copy_with_only_the_inventory_sheet(self) -> None:
        self.assertIsNotNone(create_single_sheet_workbook)
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.xlsx"
            output = Path(temp_dir) / "email.xlsx"
            workbook = Workbook()
            workbook.active.title = "Customer View"
            workbook.create_sheet("Availability")
            inventory = workbook.create_sheet("Inventory")
            inventory["A1"] = "Inventory only"
            workbook.create_sheet("Order Form")
            workbook.save(source)
            workbook.close()

            create_single_sheet_workbook(source, output, sheet_name="Inventory")

            result = load_workbook(output, data_only=False, read_only=True)
            try:
                self.assertEqual(result.sheetnames, ["Inventory"])
                self.assertEqual(result["Inventory"]["A1"].value, "Inventory only")
            finally:
                result.close()

    def test_creates_email_copy_with_static_customer_view_and_without_availability(self) -> None:
        self.assertIsNotNone(create_inventory_email_workbook)
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.xlsx"
            output = Path(temp_dir) / "email.xlsx"
            workbook = Workbook()
            customer = workbook.active
            customer.title = "Customer View"
            customer["A1"] = "Total"
            customer["A2"] = "=SUM(Availability!B2:B3)"
            customer["A2"].font = Font(bold=True)
            availability = workbook.create_sheet("Availability")
            availability.append(["Product", "Qty"])
            availability.append(["A", 5])
            availability.append(["B", 10])
            inventory = workbook.create_sheet("Inventory")
            inventory["A1"] = "Inventory"
            workbook.create_sheet("Order Form")
            workbook.save(source)
            workbook.close()

            create_inventory_email_workbook(
                source,
                output,
                customer_view_values={"A2": 15},
            )

            result = load_workbook(output, data_only=False)
            try:
                self.assertEqual(result.sheetnames, ["Customer View"])
                self.assertEqual(result["Customer View"]["A2"].value, 15)
                self.assertTrue(result["Customer View"]["A2"].font.bold)
                self.assertNotIn("Inventory", result.sheetnames)
            finally:
                result.close()

    def test_rebuilds_customer_view_from_availability_as_static_customer_matrix(self) -> None:
        self.assertIsNotNone(rebuild_customer_view_from_availability)
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.xlsx"
            workbook = Workbook()
            customer = workbook.active
            customer.title = "Customer View"
            customer["A1"] = "old formula layout"
            availability = workbook.create_sheet("Availability")
            availability.append(
                [
                    "Vendor Name",
                    "Product Description",
                    "Unit of Sale",
                    "Package Type",
                    "Pack",
                    "Units / Pack",
                    "Qty Packages",
                    "Price",
                    "Available From",
                ]
            )
            availability.append(["Pacific", "Aster Purple Bonita", "Bunch", "D", 6, 12, 2, 3.95, date(2026, 10, 9)])
            availability.append(["Pacific", "Aster Purple Bonita", "Bunch", "D", 6, 12, 3, 3.95, date(2026, 10, 10)])
            availability.append(["Pacific", "Aster Purple Bonita", "Bunch", "D", 6, 12, 0, 3.95, date(2026, 10, 11)])
            availability.append(["Pacific", "Marigold Orange", "Bunch", "L", 5, 10, 1, 3.10, date(2026, 10, 9)])
            for day in range(12, 20):
                availability.append(["Pacific", "Marigold Orange", "Bunch", "L", 5, 10, 1, 3.10, date(2026, 10, day)])
            availability.append(["Pacific", "Celosia Orange", "Bunch", "D", 10, 10, 0, 5.20, date(2026, 10, 9)])
            availability.add_table(Table(displayName="tblAvailability2", ref="A1:I14"))
            workbook.create_sheet("Inventory")
            workbook.save(source)
            workbook.close()

            variants = rebuild_customer_view_from_availability(source)

            result = load_workbook(source, data_only=False)
            try:
                customer = result["Customer View"]
                self.assertEqual(variants, 2)
                self.assertEqual(customer["A1"].value, "PACIFICA FARMS")
                self.assertEqual(customer["A1"].font.name, "The Seasons")
                self.assertEqual(customer["A1"].fill.fgColor.rgb, "FF074C73")
                self.assertEqual(customer["G8"].value.date(), date(2026, 10, 9))
                self.assertEqual(customer["G8"].fill.fgColor.rgb, "FFFFB84C")
                self.assertEqual(customer["A10"].font.name, "Lora")
                self.assertEqual(customer["H8"].value.date(), date(2026, 10, 10))
                self.assertEqual(customer["I8"].value.date(), date(2026, 10, 12))
                self.assertEqual(customer["P8"].value.date(), date(2026, 10, 19))
                self.assertEqual(customer["G10"].value, 2)
                self.assertEqual(customer["H10"].value, 3)
                self.assertEqual(customer["I10"].value, "—")
                self.assertEqual(customer["P10"].value, "—")
                self.assertEqual(customer["A12"].value, "Marigold Orange")
                self.assertEqual(customer["D9"].value, "ASTERS")
                self.assertEqual(customer.freeze_panes, "A10")
                self.assertFalse(
                    any(
                        isinstance(cell.value, str) and cell.value.startswith("=")
                        for row in customer.iter_rows()
                        for cell in row
                    )
                )
            finally:
                result.close()

    def test_rebuilds_customer_view_with_explicit_greens_and_mixed_boxes_categories(self) -> None:
        self.assertIsNotNone(rebuild_customer_view_from_availability)
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.xlsx"
            workbook = Workbook()
            customer = workbook.active
            customer.title = "Customer View"
            availability = workbook.create_sheet("Availability")
            availability.append(
                [
                    "Vendor Name",
                    "Product Description",
                    "Unit of Sale",
                    "Package Type",
                    "Pack",
                    "Units / Pack",
                    "Qty Packages",
                    "Price",
                    "Available From",
                ]
            )
            # Deliberately interleave the categories in Availability.
            availability.append(["Pacific", "Day of Dead Sampler", "Bunch", "L", 10, 80, 1, 39.5, date(2026, 10, 9)])
            availability.append(["Pacific", "Myrtle Green 60cm", "Bunch", "L", 10, 10, 1, 4.0, date(2026, 10, 9)])
            availability.append(["Pacific", "Greens Sampler", "Bunch", "L", 10, 10, 1, 4.0, date(2026, 10, 9)])
            availability.append(["Pacific", "Parvifolia Green 50cm", "Bunch", "L", 10, 10, 1, 4.0, date(2026, 10, 9)])
            availability.add_table(Table(displayName="tblAvailability2", ref="A1:I5"))
            workbook.save(source)
            workbook.close()

            variants = rebuild_customer_view_from_availability(source)

            result = load_workbook(source, data_only=False)
            try:
                customer = result["Customer View"]
                self.assertEqual(variants, 4)
                category_rows = {
                    customer.cell(row=row, column=4).value: row
                    for row in range(9, customer.max_row + 1)
                    if customer.cell(row=row, column=4).value in {"GREENS", "MIXED BOXES"}
                }
                self.assertEqual(list(category_rows), ["GREENS", "MIXED BOXES"])
                greens_row = category_rows["GREENS"]
                mixed_row = category_rows["MIXED BOXES"]
                self.assertEqual(
                    [customer.cell(row=row, column=1).value for row in range(greens_row + 1, mixed_row)],
                    ["Greens Sampler", "Myrtle Green 60cm", "Parvifolia Green 50cm"],
                )
                self.assertEqual(customer.cell(row=mixed_row + 1, column=1).value, "Day of Dead Sampler")
            finally:
                result.close()

    def test_keeps_the_second_sheet_in_the_full_transformed_workbook(self) -> None:
        self.assertIsNotNone(transform_inventory_workbook)
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.xlsx"
            output = Path(temp_dir) / "output.xlsx"
            workbook = Workbook()
            active = workbook.active
            active.title = "Customer View"
            active["A1"] = "Debe conservarse"
            availability = workbook.create_sheet("Availability")
            availability.append(["Product Description", "Available From"])
            availability.append(["A", date(2026, 9, 14)])
            workbook.save(source)
            workbook.close()

            transform_inventory_workbook(source, output, assumed_today=self.assumed_today)

            result = load_workbook(output, data_only=False)
            try:
                self.assertEqual(result.sheetnames, ["Customer View", "Availability"])
                self.assertEqual(result["Customer View"]["A1"].value, "Debe conservarse")
                self.assertEqual(result["Availability"]["A3"].value, "A")
            finally:
                result.close()

    def test_verifies_added_rows_against_their_retained_source_rows(self) -> None:
        self.assertIsNotNone(transform_inventory_workbook)
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.xlsx"
            output = Path(temp_dir) / "output.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Availability"
            sheet.append(["Product Description", "Available From"])
            sheet.append(["A", date(2026, 9, 14)])
            sheet.append(["B", date(2026, 9, 10)])
            sheet.append(["B", date(2026, 9, 15)])
            sheet["A2"].fill = PatternFill(fill_type="solid", fgColor="FF0000")
            sheet["A4"].fill = PatternFill(fill_type="solid", fgColor="00FF00")
            workbook.save(source)
            workbook.close()

            transform_inventory_workbook(source, output, assumed_today=self.assumed_today)

            result = load_workbook(output, data_only=False)
            try:
                self.assertEqual(result.active["A4"].value, "A")
                self.assertEqual(result.active["A5"].value, "B")
                self.assertEqual(result.active["A5"]._style, result.active["A3"]._style)
            finally:
                result.close()

    def test_omits_customer_view_variants_with_no_positive_quantities(self) -> None:
        self.assertIsNotNone(rebuild_customer_view_from_availability)
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.xlsx"
            workbook = Workbook()
            customer = workbook.active
            customer.title = "Customer View"
            availability = workbook.create_sheet("Availability")
            availability.append(
                [
                    "Vendor Name",
                    "Product Description",
                    "Unit of Sale",
                    "Package Type",
                    "Pack",
                    "Units / Pack",
                    "Qty Packages",
                    "Price",
                    "Available From",
                ]
            )
            availability.append(["Pacific", "Marigold Orange", "Bunch", "L", 5, 10, 2, 3.10, date(2026, 10, 9)])
            availability.append(["Pacific", "Celosia Orange", "Bunch", "D", 10, 10, 0, 5.20, date(2026, 10, 9)])
            availability.add_table(Table(displayName="tblAvailability2", ref="A1:I3"))
            workbook.save(source)
            workbook.close()

            variants = rebuild_customer_view_from_availability(source)

            result = load_workbook(source, data_only=False)
            try:
                customer = result["Customer View"]
                self.assertEqual(variants, 1)
                values = [
                    customer.cell(row=row, column=1).value
                    for row in range(1, customer.max_row + 1)
                ]
                self.assertIn("Marigold Orange", values)
                self.assertNotIn("Celosia Orange", values)
                self.assertEqual(customer["G8"].value.date(), date(2026, 10, 9))
                self.assertIsNone(customer["H8"].value)
                self.assertIsNone(customer["H8"].fill.fill_type)
                self.assertIsNone(customer["H9"].fill.fill_type)
            finally:
                result.close()

    def test_ignores_formatted_blank_columns_after_availability_data(self) -> None:
        self.assertIsNotNone(transform_inventory_workbook)
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.xlsx"
            output = Path(temp_dir) / "output.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Availability"
            sheet.append(["Product Description", "Available From"])
            sheet.append(["A", date(2026, 9, 14)])
            sheet["D2"].fill = PatternFill(fill_type="solid", fgColor="FFFF00")
            sheet["D3"].fill = PatternFill(fill_type="solid", fgColor="00FF00")
            workbook.save(source)
            workbook.close()

            transform_inventory_workbook(
                source,
                output,
                assumed_today=self.assumed_today,
            )

            result = load_workbook(output, data_only=False)
            try:
                self.assertEqual(result["Availability"]["A3"].value, "A")
                self.assertEqual(result["Availability"]["B3"].value.date(), date(2026, 9, 15))
            finally:
                result.close()

    def test_replaces_inventory_and_updates_availability_without_increasing_qty(self) -> None:
        self.assertIsNotNone(refresh_workbook_with_komet_inventory)
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.xlsx"
            komet = Path(temp_dir) / "komet.xlsx"
            output = Path(temp_dir) / "output.xlsx"

            workbook = Workbook()
            customer = workbook.active
            customer.title = "Customer View"
            customer["A1"] = "=SUM(tblAvailability2[Qty Packages])"
            availability = workbook.create_sheet("Availability")
            availability.append(
                [
                    "Vendor Name",
                    "Product Description",
                    "Unit of Sale",
                    "Package Type",
                    "Pack",
                    "Units / Pack",
                    "Qty Packages",
                    "Price",
                    "Available From",
                ]
            )
            availability.append(["Pacific", "A product", "Bunch", "L", 10, 1, "=OLD", 3, date(2026, 9, 15)])
            availability.append(["Pacific", "B product", "Bunch", "L", 10, 1, 3, 3, date(2026, 9, 15)])
            availability.append(["Pacific", "Missing product", "Bunch", "L", 10, 1, "=OLD", 3, date(2026, 9, 15)])
            availability.add_table(Table(displayName="tblAvailability2", ref="A1:I4"))

            inventory = workbook.create_sheet("Inventory")
            inventory.append(["Inventory source"])
            for _ in range(6):
                inventory.append([])
            inventory_headers = ["AWB", "Ref #", "Location", "Product", "Hold", "Customer", "Vendor", "Aging", "Qty"]
            inventory.append(inventory_headers)
            inventory.append(["AWB-2026-09-14", "old", "L", "A product", "", "", "", 1, 8])
            inventory.append(["AWB-2026-09-14", "old", "L", "Missing product", "", "", "", 1, 4])
            inventory.add_table(Table(displayName="tblInventory", ref="A8:I10"))
            workbook.save(source)
            workbook.close()

            export = Workbook()
            export_sheet = export.active
            export_sheet.append(inventory_headers)
            export_sheet.append(["AWB-2026-09-10", "new", "L", "A product", "", "", "", 5, 5])
            export_sheet.append(["AWB-2026-09-10", "new", "L", "B product", "", "", "", 5, 9])
            export.save(komet)
            export.close()

            result = refresh_workbook_with_komet_inventory(
                source,
                komet,
                output,
                assumed_today=self.assumed_today,
            )

            workbook = load_workbook(output, data_only=False)
            try:
                self.assertEqual(workbook.sheetnames, ["Customer View", "Availability", "Inventory"])
                self.assertEqual(workbook["Customer View"]["A1"].value, "=SUM(tblAvailability2[Qty Packages])")
                self.assertEqual([workbook["Availability"][f"G{row}"].value for row in (2, 3, 4)], [5, 3, 0])
                self.assertTrue(all(not str(workbook["Availability"][f"G{row}"].value).startswith("=") for row in (2, 3, 4)))
                self.assertEqual(workbook["Inventory"]["D9"].value, "A product")
                self.assertEqual(workbook["Inventory"]["I9"].value, 5)
                self.assertEqual(workbook["Inventory"].tables["tblInventory"].ref, "A8:I10")
            finally:
                workbook.close()

            self.assertEqual(result.updated_rows, 2)
            self.assertEqual(result.decreased_rows, 2)
            self.assertEqual(result.before_total, 15)
            self.assertEqual(result.after_total, 8)

            preserved_output = Path(temp_dir) / "output-without-komet-export.xlsx"
            preserved_result = refresh_workbook_with_komet_inventory(
                source,
                None,
                preserved_output,
                assumed_today=self.assumed_today,
            )
            preserved = load_workbook(preserved_output, data_only=False)
            try:
                self.assertEqual([preserved["Availability"][f"G{row}"].value for row in (2, 3, 4)], [8, 3, 4])
                self.assertEqual(preserved["Inventory"]["I9"].value, 8)
                self.assertEqual(preserved["Inventory"].tables["tblInventory"].ref, "A8:I10")
            finally:
                preserved.close()
            self.assertEqual(preserved_result.inventory_rows, 2)
            self.assertEqual(preserved_result.updated_rows, 0)
            self.assertEqual(preserved_result.decreased_rows, 0)
            self.assertEqual(preserved_result.before_total, 15)
            self.assertEqual(preserved_result.after_total, 15)


if __name__ == "__main__":
    unittest.main()
