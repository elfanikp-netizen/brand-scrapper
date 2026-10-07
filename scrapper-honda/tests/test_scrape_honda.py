import unittest
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import Mock, patch

from openpyxl import Workbook, load_workbook

from scrape_honda import (
    load_results,
    normalize_part_number,
    parse_args,
    parse_msrp,
    prepare_crawl_rows,
    product_page_matches_oem,
    product_url_matches_oem,
    product_status,
    save_results,
    status_is_complete,
    wait_for_human_verification,
)


class HondaCrawlerTests(unittest.TestCase):
    def test_human_verification_waits_after_request_failure_without_reloading(self):
        page = Mock()
        failures = ["challenge request failed: net::ERR_NAME_NOT_RESOLVED"]
        with patch("scrape_honda.page_needs_human_verification", side_effect=[True, True, False]):
            wait_for_human_verification(page, failures)
        page.reload.assert_not_called()
        self.assertEqual(page.wait_for_timeout.call_count, 1)

    def test_search_retries_are_opt_in(self):
        with patch("sys.argv", ["scrape_honda.py", "--input", "parts.xlsx", "--column", "A"]):
            self.assertFalse(parse_args().attempt_on_search)
        with patch(
            "sys.argv",
            ["scrape_honda.py", "--input", "parts.xlsx", "--column", "A", "--attempt-on-search"],
        ):
            self.assertTrue(parse_args().attempt_on_search)

    def test_normalize_part_number(self):
        self.assertEqual(normalize_part_number("17220-R5A-A00"), "17220R5AA00")

    def test_parse_msrp_uses_labelled_original_price_not_discount(self):
        self.assertEqual(parse_msrp("$21.76 MSRP: $31.00\nYou Save: $9.24"), 31.00)
        self.assertEqual(parse_msrp("$23.58\nSave\n$7.42\n23.9% off\nMSRP $31.00"), 31.00)
        self.assertEqual(parse_msrp("MSRP $31.00"), 31.00)
        self.assertIsNone(parse_msrp("$21.76 Available from multiple sellers"))

    def test_product_url_matches_requested_honda_oem(self):
        self.assertTrue(
            product_url_matches_oem(
                "https://honda.oempartsonline.com/oem-parts/honda-engine-air-filter-17220r5aa00",
                "17220-R5A-A00",
            )
        )
        self.assertFalse(
            product_url_matches_oem(
                "https://honda.oempartsonline.com/oem-parts/honda-engine-air-filter-17220r5aa00",
                "17220-R5A-A01",
            )
        )

    def test_product_page_matches_when_url_slug_drops_leading_zero(self):
        self.assertTrue(
            product_page_matches_oem(
                "https://honda.oempartsonline.com/oem-parts/honda-bumper-cover-4711snaa90zz",
                "Bumper Cover - Honda (04711-SNA-A90ZZ)\nPart Number (MPN)\n04711-SNA-A90ZZ",
                "04711SNAA90ZZ",
            )
        )

    def test_status_identifies_superseded_manufacturer_part_number(self):
        self.assertEqual(
            product_status(
                "Manufacturer Part Number\n17220-R5A-A01", "17220-R5A-A00", 31.00
            ),
            "success_superseded: 17220-R5A-A01",
        )

    def test_only_success_statuses_complete(self):
        self.assertTrue(status_is_complete("success"))
        self.assertTrue(status_is_complete("success_superseded: 17220-R5A-A01"))
        self.assertFalse(status_is_complete("not_found"))
        self.assertFalse(status_is_complete("msrp_not_found"))

    def test_duplicate_oems_are_queued_as_separate_input_rows(self):
        results, pending = prepare_crawl_rows(["17220-R5A-A00", "17220-R5A-A00"], [])
        self.assertEqual(len(results), 2)
        self.assertEqual(pending, [(0, "17220-R5A-A00"), (1, "17220-R5A-A00")])

    def test_resume_matches_duplicate_results_by_occurrence(self):
        saved = [
            ("17220-R5A-A00", 31.00, date(2026, 10, 1), "success"),
            ("17220-R5A-A00", None, date(2026, 10, 1), "not_found"),
        ]
        results, pending = prepare_crawl_rows(["17220-R5A-A00", "17220-R5A-A00"], saved)
        self.assertEqual(results[0][1:], saved[0][1:])
        self.assertEqual(pending, [(1, "17220-R5A-A00")])

    def test_output_status_round_trip(self):
        with TemporaryDirectory() as directory:
            output = Path(directory) / "results.xlsx"
            rows = [("17220-R5A-A00", 31.00, date(2026, 10, 1), "success_superseded: 17220-R5A-A01")]
            save_results(output, rows)
            workbook = load_workbook(output, data_only=True)
            self.assertEqual(tuple(cell.value for cell in workbook.active[1]), ("OEM Number", "MSRP", "Date", "Status"))
            workbook.close()
            self.assertEqual(load_results(output)[0][3], "success_superseded: 17220-R5A-A01")


if __name__ == "__main__":
    unittest.main()