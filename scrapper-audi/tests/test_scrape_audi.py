import unittest
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import Mock, patch

from openpyxl import Workbook, load_workbook

from scrape_audi import (
    crawl_msrp,
    detail_status,
    detail_url_matches_oem,
    is_audi_site_url,
    is_product_detail_page,
    load_results,
    normalize_part_number,
    parse_args,
    parse_msrp,
    parse_product_msrp,
    prepare_crawl_rows,
    read_oem_values,
    resolve_column,
    save_results,
    status_is_complete,
    wait_for_human_verification,
)


class AudiCrawlerTests(unittest.TestCase):
    def test_parse_args_accepts_standard_options(self):
        with patch("sys.argv", ["scrape_audi.py", "--input", "parts.xlsx", "--column", "A"]):
            args = parse_args()
        self.assertEqual(args.input, Path("parts.xlsx"))
        self.assertEqual(args.column, "A")

    def test_normalize_part_number(self):
        self.assertEqual(normalize_part_number("06L-115-562-B"), "06L115562B")

    def test_parse_msrp_returns_original_price_not_discount(self):
        self.assertEqual(parse_msrp("$12.86 MSRP: $17.15\\nYou Save: $4.29"), 17.15)

    def test_parse_product_msrp_from_audi_detail_text(self):
        text = """Audi 06L-115-562-B Filter Element
2014-2025 Audi 06L115562B
Part Description
Filterelem
$12.86 MSRP: $17.15
Ships in 1-2 Business Days"""
        self.assertEqual(parse_product_msrp(text), 17.15)

    def test_live_audi_oem_msrp(self):
        live_price_text = "$12.86 MSRP: $17.15\\nYou Save: $4.29 (26%)\\nShips in 1-2 Business Days"
        self.assertEqual(parse_msrp(live_price_text), 17.15)

    def test_parse_product_msrp_requires_audi_part_reference(self):
        self.assertIsNone(parse_product_msrp("MSRP: $17.15"))

    def test_detail_url_matches_audi_genuine_route_without_html_suffix(self):
        url = "https://www.audipartsgiant.com/genuine/audi~filter-element~06l115562b"
        self.assertTrue(detail_url_matches_oem(url, "06L-115-562-B"))
        self.assertFalse(detail_url_matches_oem(url, "06L-115-562-C"))
        self.assertFalse(detail_url_matches_oem("https://example.com/genuine/audi~filter~06l115562b", "06L-115-562-B"))

    def test_site_url_recognizes_home_and_genuine_page(self):
        self.assertTrue(is_audi_site_url("https://www.audipartsgiant.com/"))
        self.assertTrue(is_audi_site_url("https://www.audipartsgiant.com/genuine/audi~filter-element~06l115562b"))
        self.assertFalse(is_audi_site_url("about:blank"))

    def test_direct_genuine_route_is_a_product_detail_page(self):
        page = Mock()
        page.url = "https://www.audipartsgiant.com/genuine/audi~filter-element~06l115562b"
        self.assertTrue(is_product_detail_page(page))

    def test_status_identifies_superseded_audi_part(self):
        self.assertEqual(
            detail_status("Manufacturer Part Number\n06L-115-562-B", "06L-115-562", 17.15),
            "success_superseded: 06L-115-562-B",
        )
        self.assertEqual(detail_status("Audi part\nMSRP: $17.15", "06L-115-562-B", 17.15), "success")

    def test_only_success_statuses_are_complete(self):
        self.assertTrue(status_is_complete("success"))
        self.assertTrue(status_is_complete("success_superseded: 06L-115-562-B"))
        self.assertFalse(status_is_complete("not_found"))
        self.assertFalse(status_is_complete("msrp_not_found"))

    def test_missing_result_returns_not_found_after_single_search(self):
        page = Mock()
        page.url = "https://www.audipartsgiant.com/"
        search_box = Mock()
        with (
            patch("scrape_audi.dismiss_blocking_popups"),
            patch("scrape_audi.find_search_box", return_value=search_box),
            patch("scrape_audi.wait_for_matching_result", return_value=None),
            patch("scrape_audi.wait_for_human_verification"),
            patch("scrape_audi.is_product_detail_page", return_value=False),
        ):
            result = crawl_msrp(page, "06H115403", [])
        self.assertEqual(result, (None, "not_found"))
        search_box.fill.assert_called_once_with("06H115403")
        search_box.press.assert_called_once_with("Enter")

    def test_direct_search_navigation_extracts_audi_msrp(self):
        page = Mock()
        page.url = "https://www.audipartsgiant.com/"
        search_box = Mock()
        search_box.press.side_effect = lambda *_: setattr(
            page, "url", "https://www.audipartsgiant.com/genuine/audi~filter-element~06l115562b"
        )
        price_text = "$12.86 MSRP: $17.15\\nYou Save: $4.29 (26%)"
        with (
            patch("scrape_audi.dismiss_blocking_popups") as dismiss_popups,
            patch("scrape_audi.find_search_box", return_value=search_box),
            patch("scrape_audi.wait_for_human_verification"),
            patch("scrape_audi.wait_for_product_msrp", return_value=price_text),
        ):
            result = crawl_msrp(page, "06L115562B", [])
        self.assertEqual(result, (17.15, "success"))
        page.wait_for_url.assert_called_once()
        self.assertGreaterEqual(dismiss_popups.call_count, 3)

    def test_human_verification_waits_without_reloading(self):
        page = Mock()
        with patch("scrape_audi.page_needs_human_verification", side_effect=[True, True, False]):
            wait_for_human_verification(page, ["challenge request failed"])
        page.reload.assert_not_called()
        page.wait_for_timeout.assert_called_once_with(1_000)

    def test_resolve_column_supports_header_and_excel_letter(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(("Description", "OEM Number"))
        sheet.append(("Oil filter", "06L115562B"))
        self.assertEqual(resolve_column(sheet, "OEM Number", 1), (2, 2))
        self.assertEqual(resolve_column(sheet, "B", 1), (2, 2))
        workbook.close()

    def test_read_oem_values(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "parts.xlsx"
            workbook = Workbook()
            workbook.active.append(("Description", "OEM Number"))
            workbook.active.append(("Oil filter", "06L115562B"))
            workbook.save(path)
            workbook.close()
            self.assertEqual(read_oem_values(path, None, "OEM Number", 1), ["06L115562B"])

    def test_duplicate_oems_resume_by_occurrence(self):
        saved = [
            ("06L115562B", 17.15, date(2026, 10, 9), "success"),
            ("06L115562B", None, date(2026, 10, 9), "not_found"),
        ]
        results, pending = prepare_crawl_rows(["06L115562B", "06L115562B"], saved)
        self.assertEqual(results[0][1:], saved[0][1:])
        self.assertEqual(pending, [(1, "06L115562B")])

    def test_success_without_msrp_is_recrawled(self):
        saved = [("06L115562B", None, date(2026, 10, 9), "success")]
        results, pending = prepare_crawl_rows(["06L115562B"], saved)
        self.assertIsNone(results[0][1])
        self.assertEqual(pending, [(0, "06L115562B")])

    def test_save_and_load_results_round_trip(self):
        with TemporaryDirectory() as directory:
            output = Path(directory) / "results.xlsx"
            rows = [("06L115562B", 17.15, date(2026, 10, 9), "success")]
            save_results(output, rows)
            self.assertEqual(load_results(output)[0][:2], ("06L115562B", 17.15))
            workbook = load_workbook(output, data_only=True)
            self.assertEqual(tuple(cell.value for cell in workbook.active[1]), ("OEM Number", "MSRP", "Date", "Status"))
            workbook.close()


if __name__ == "__main__":
    unittest.main()
