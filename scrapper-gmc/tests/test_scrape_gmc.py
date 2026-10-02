import unittest
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook

from scrape_gmc import (
    load_results,
    normalize_part_number,
    parse_msrp,
    prepare_crawl_rows,
    product_status,
    product_url_matches_oem,
    save_results,
    status_is_complete,
)


class GmcCrawlerTests(unittest.TestCase):
    def test_normalize_part_number(self):
        self.assertEqual(normalize_part_number("126-90512"), "12690512")

    def test_parse_msrp_uses_labelled_price_not_other_amounts(self):
        self.assertEqual(parse_msrp("Starting at\n$51.92\nMSRP\n$51.92*"), 51.92)
        self.assertIsNone(parse_msrp("$46.73 Available from multiple sellers"))

    def test_product_url_matches_queried_replacement_oem(self):
        self.assertTrue(
            product_url_matches_oem(
                "https://parts.gmparts.com/product/evaporative-valve-12737252", "12737252"
            )
        )
        self.assertFalse(
            product_url_matches_oem(
                "https://parts.gmparts.com/product/evaporative-valve-12737252", "12690512"
            )
        )

    def test_status_identifies_superseded_part_from_card(self):
        self.assertEqual(
            product_status("GM Part #12737252", "GM Part #12737252\n(Replaces: #12690512)", "12690512", 51.92),
            "success_superseded: 12737252",
        )

    def test_only_success_statuses_complete(self):
        self.assertTrue(status_is_complete("success"))
        self.assertTrue(status_is_complete("success_superseded: 12737252"))
        self.assertFalse(status_is_complete("not_found"))
        self.assertFalse(status_is_complete("msrp_not_found"))

    def test_duplicate_oems_are_queued_as_separate_input_rows(self):
        results, pending = prepare_crawl_rows(["12690512", "12690512"], [])
        self.assertEqual(len(results), 2)
        self.assertEqual(pending, [(0, "12690512"), (1, "12690512")])

    def test_resume_matches_duplicate_results_by_occurrence(self):
        saved = [
            ("12690512", 51.92, date(2026, 10, 1), "success"),
            ("12690512", None, date(2026, 10, 1), "not_found"),
        ]
        results, pending = prepare_crawl_rows(["12690512", "12690512"], saved)
        self.assertEqual(results[0][1:], saved[0][1:])
        self.assertEqual(pending, [(1, "12690512")])

    def test_output_status_round_trip(self):
        with TemporaryDirectory() as directory:
            output = Path(directory) / "results.xlsx"
            rows = [("12690512", 51.92, date(2026, 10, 1), "success_superseded: 12737252")]
            save_results(output, rows)
            workbook = load_workbook(output, data_only=True)
            self.assertEqual(tuple(cell.value for cell in workbook.active[1]), ("OEM Number", "MSRP", "Date", "Status"))
            workbook.close()
            self.assertEqual(load_results(output)[0][3], "success_superseded: 12737252")


if __name__ == "__main__":
    unittest.main()