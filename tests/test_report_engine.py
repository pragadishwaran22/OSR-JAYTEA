from __future__ import annotations

import unittest
from io import BytesIO

from openpyxl import Workbook, load_workbook

from report_engine import (
    ReportInputError,
    generate_workbook,
    prepare_report,
    validate_generated_workbook,
)


def sample_source() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Order Status-By shipment Date"
    headers = [
        "Ord Id", "Buyer Id", "Buyer Name", "Buyer Requested Shipment Date",
        "Prod Id", "Prod Name", "Pending Prod", "Vsl Saling Date",
        "Planned in 37", "Planned in 38", "Planned in 39",
        "Planned Qty Total", "Un Planned",
    ]
    for column, header in enumerate(headers, start=1):
        sheet.cell(6, column, header)

    source_row = 9
    for order_number in range(1, 38):
        if order_number == 2:
            product_names = ["TEST PRODUCT STD"]
        elif order_number == 1:
            product_names = ["TEST PRODUCT 1", "TEST PRODUCT STD MIXED"]
        else:
            product_names = [f"TEST PRODUCT {order_number}"]
        for item_number, product_name in enumerate(product_names):
            pending = 100
            if order_number == 1:
                planned = [0, 0, 0]
            elif order_number == 3:
                planned = [10, 0, 20]
            elif order_number == 4:
                planned = [10, 0, 90]
            elif order_number == 5:
                planned = [30, 0, 0]
            else:
                planned = [100, 0, 0]
            values = [
                f"ORD-{order_number}" if item_number == 0 else None,
                "BUYER-1" if item_number == 0 else None,
                "Sample buyer" if item_number == 0 else None,
                "2026-09-01" if item_number == 0 else None,
                f"FG-{order_number}-{item_number}",
                product_name,
                pending,
                "2026-09-05",
                *planned,
                sum(planned),
                max(pending - sum(planned), 0),
            ]
            for column, value in enumerate(values, start=1):
                sheet.cell(source_row, column, value)
            source_row += 1
    buffer = BytesIO()
    workbook.save(buffer)
    workbook.close()
    return buffer.getvalue()


class ReportEngineTests(unittest.TestCase):
    def test_complete_workbook_reconciles(self):
        prepared = prepare_report(sample_source(), "daily.xlsx")
        self.assertEqual(36, len(prepared["eligible_order_ids"]))
        self.assertEqual(36, len(prepared["rows"]))
        self.assertEqual(2, len(prepared["excluded_std"]))
        self.assertEqual(["ORD-2"], prepared["all_std_order_ids"])
        self.assertEqual(1, prepared["condition_counts"]["Red"])
        self.assertEqual(1, prepared["condition_counts"]["Black"])
        self.assertEqual(1, prepared["condition_counts"]["Yellow"])
        self.assertEqual(1, prepared["condition_counts"]["Green"])

        output = generate_workbook(prepared, "all")
        validate_generated_workbook(output, prepared, "all")
        workbook = load_workbook(BytesIO(output), data_only=False)
        self.assertEqual(
            [
                "Highlighted Report", "Red Only", "Black Only", "Yellow Only",
                "Green Only", "Excluded STD Log", "Run Summary",
            ],
            workbook.sheetnames,
        )
        headers = [
            workbook["Highlighted Report"].cell(5, column).value
            for column in range(1, workbook["Highlighted Report"].max_column + 1)
        ]
        self.assertNotIn("Vsl Saling Date", headers)
        workbook.close()

    def test_condition_only_workbook(self):
        prepared = prepare_report(sample_source(), "daily.xlsx")
        output = generate_workbook(prepared, "red")
        validate_generated_workbook(output, prepared, "red")
        workbook = load_workbook(BytesIO(output), data_only=False)
        self.assertEqual(["Red Only"], workbook.sheetnames)
        self.assertEqual(6, workbook["Red Only"].max_row)
        workbook.close()

    def test_missing_required_sheet_stops(self):
        workbook = Workbook()
        buffer = BytesIO()
        workbook.save(buffer)
        workbook.close()
        with self.assertRaisesRegex(ReportInputError, "Required sheet not found"):
            prepare_report(buffer.getvalue(), "daily.xlsx")


if __name__ == "__main__":
    unittest.main()
