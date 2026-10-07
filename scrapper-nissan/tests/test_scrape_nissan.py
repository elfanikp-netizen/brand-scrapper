import unittest
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import Mock, patch

from openpyxl import Workbook, load_workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from scrape_nissan import (
    crawl_msrp,
    detail_status,
    detail_url_matches_oem,
    dismiss_blocking_popups,
    find_matching_result,
    is_matching_detail_page,
    is_nissan_site_url,
    load_results,
    normalize_part_number,
    parse_args,
    parse_msrp,
    parse_product_msrp,
    prepare_crawl_rows,
    read_oem_values,
    resolve_column,
    save_results,
    should_retry_crawl,
    status_is_complete,
    wait_for_human_verification,
)


class NissanCrawlerTests(unittest.TestCase):
    def test_full_search_is_disabled_by_default(self):
        with patch("sys.argv", ["scrape_nissan.py", "--input", "parts.xlsx", "--column", "A"]):
            self.assertFalse(parse_args().attempt_on_search)
        with patch(
            "sys.argv",
            ["scrape_nissan.py", "--input", "parts.xlsx", "--column", "A", "--attempt-on-search"],
        ):
            self.assertTrue(parse_args().attempt_on_search)

    def test_missing_suggestion_submits_search_once(self):
        page = Mock()
        search_box = Mock()
        page.title.return_value = "Nissan Parts"
        page.url = "https://www.nissanpartsdeal.com/"
        page.locator.return_value.inner_text.return_value = "Nissan Parts"
        with (
            patch("scrape_nissan.dismiss_blocking_popups"),
            patch("scrape_nissan.find_search_box", return_value=search_box),
            patch("scrape_nissan.wait_for_matching_result", return_value=None),
            patch("scrape_nissan.page_needs_human_verification", return_value=False),
        ):
            result = crawl_msrp(page, "12030-2N100", [])
        self.assertEqual(result, (None, "not_found"))
        search_box.fill.assert_called_once_with("12030-2N100")
        self.assertEqual(search_box.press.call_count, 1)

    def test_search_uses_enter_then_clicks_button_when_needed(self):
        page = Mock()
        search_box = Mock()
        search_button = Mock()
        page.title.return_value = "Nissan Parts"
        page.url = "https://www.nissanpartsdeal.com/"
        page.locator.return_value.inner_text.return_value = "Nissan Parts"
        page.wait_for_url.side_effect = PlaywrightTimeoutError("no navigation")
        with (
            patch("scrape_nissan.dismiss_blocking_popups"),
            patch("scrape_nissan.find_search_box", return_value=search_box),
            patch("scrape_nissan.find_search_button", return_value=search_button),
            patch("scrape_nissan.is_product_detail_page", return_value=True),
            patch("scrape_nissan.wait_for_product_detail", return_value=True),
            patch("scrape_nissan.is_matching_detail_page", return_value=True),
            patch("scrape_nissan.wait_for_product_msrp", return_value="Nissan 12030-2N100\nMSRP: $30.06"),
            patch("scrape_nissan.parse_product_msrp", return_value=30.06),
            patch("scrape_nissan.detail_status", return_value="success"),
        ):
            result = crawl_msrp(page, "12030-2N100", [])
        self.assertEqual(result, (30.06, "success"))
        search_box.fill.assert_called_once_with("12030-2N100")
        search_box.press.assert_called_once_with("Enter")
        search_button.click.assert_called_once_with(timeout=10_000)

    def test_verification_waits_for_manual_completion_after_request_failure(self):
        page = Mock()
        failures = ["challenge request failed: net::ERR_NAME_NOT_RESOLVED"]
        with patch("scrape_nissan.page_needs_human_verification", side_effect=[True, True, False]):
            wait_for_human_verification(page, failures)
        page.reload.assert_not_called()
        self.assertEqual(page.wait_for_timeout.call_count, 1)

    def test_dismiss_blocking_popups_closes_icon_modal(self):
        page = Mock()
        close_button = Mock()
        close_button.is_visible.return_value = True
        close_button.is_enabled.return_value = True
        page.locator.return_value.count.return_value = 1
        page.locator.return_value.nth.return_value = close_button

        dismiss_blocking_popups(page)

        page.locator.assert_called_with(
            'button[aria-label*="close" i], button[title*="close" i], '
            'button[aria-label*="dismiss" i], button:has-text("Close"), '
            'button:has-text("Dismiss"), button:has-text("Not now"), button:has-text("No thanks"), '
            'button:has-text("×"), [class*="close" i]:has-text("×"), '
            '.ab-modal *:has-text("×")'
        )
        self.assertGreaterEqual(close_button.click.call_count, 1)

    def test_normalize_part_number(self):
        self.assertEqual(normalize_part_number("12030-2N100"), "120302N100")

    def test_parse_msrp_with_currency_and_thousands_separator(self):
        self.assertEqual(parse_msrp("Nissan 12030-2N100\nMSRP: $1,030.06"), 1030.06)

    def test_parse_product_price_uses_original_msrp_not_discount(self):
        text = """Nissan 12030-2N100 Air Cleaner Filter
Part Description
Filter-Air Cleaner
$23.26 MSRP: $30.06
You Save: $6.80 (23%)
Related Parts
Nissan 12030-2N101
MSRP $999.99"""
        self.assertEqual(parse_product_msrp(text), 30.06)

    def test_parse_product_price_falls_back_to_msrp_without_sale_price(self):
        self.assertEqual(parse_product_msrp("Nissan 12030-2N100 Air Cleaner\nMSRP $30.06"), 30.06)

    def test_parse_product_price_accepts_msrp_when_no_nissan_heading_is_present(self):
        text = "Manufacturer Part Number\n12030-2N100\nMSRP: $30.06"
        self.assertEqual(parse_product_msrp(text), 30.06)

    def test_parse_product_price_accepts_msrp_price_label_when_no_nissan_heading_is_present(self):
        text = "Manufacturer Part Number\nH01004AFMB\nMSRP Price: $30.06"
        self.assertEqual(parse_product_msrp(text), 30.06)

    def test_parse_product_price_accepts_pn_price_wrap_content(self):
        text = "Manufacturer Part Number\nH01004AFMB\nMSRP Price\n$30.06"
        self.assertEqual(parse_product_msrp(text), 30.06)

    def test_parse_product_price_requires_nissan_part_reference(self):
        self.assertIsNone(parse_product_msrp("MSRP $30.06"))

    def test_find_matching_result_accepts_hidden_product_links(self):
        page = Mock()
        link = Mock()
        link.is_visible.return_value = False
        link.inner_text.return_value = "Nissan 12030-2N100 Air Cleaner"
        link.get_attribute.return_value = "/genuine/nissan-air-filter~12030-2n100.html"
        page.locator.return_value.count.return_value = 1
        page.locator.return_value.nth.return_value = link

        self.assertIsNotNone(find_matching_result(page, "12030-2N100"))

    def test_matching_detail_page_accepts_visible_oem_without_nissan_label(self):
        page = Mock()
        page.url = "https://www.nissanpartsdeal.com/genuine/nissan-air-filter~12030-2n100.html"
        page.locator.return_value.inner_text.return_value = "12030-2N100\nMSRP: $30.06"

        self.assertTrue(is_matching_detail_page(page, "12030-2N100"))

    def test_detail_url_matches_nissan_part_url(self):
        genuine_url = "https://www.nissanpartsdeal.com/genuine/nissan-air-filter~12030-2n100.html"
        parts_url = "https://www.nissanpartsdeal.com/parts/nissan-door-fr-rh~h0100-4afmb.html"
        self.assertTrue(detail_url_matches_oem(genuine_url, "12030-2N100"))
        self.assertTrue(detail_url_matches_oem(parts_url, "H01004AFMB"))
        self.assertFalse(detail_url_matches_oem(genuine_url, "12030-2N101"))

    def test_site_url_recognizes_nissan_home_and_product_pages(self):
        self.assertTrue(is_nissan_site_url("https://www.nissanpartsdeal.com/"))
        self.assertTrue(
            is_nissan_site_url(
                "https://www.nissanpartsdeal.com/genuine/nissan-air-filter~12030-2n100.html"
            )
        )
        self.assertTrue(is_nissan_site_url("https://www.nissanpartsdeal.com/parts/nissan-door-fr-rh~h0100-4afmb.html"))
        self.assertFalse(is_nissan_site_url("about:blank"))

    def test_status_reports_superseded_nissan_part(self):
        self.assertEqual(
            detail_status("Nissan 12030-2N101 Air Cleaner Filter", "12030-2N100", 30.06),
            "success_superseded: 12030-2N101",
        )
        self.assertEqual(detail_status("Nissan Parts\nMSRP $30.06", "12030-2N100", 30.06), "success")

    def test_only_successful_statuses_are_complete(self):
        self.assertTrue(status_is_complete("success"))
        self.assertTrue(status_is_complete("success_superseded: 12030-2N101"))
        self.assertFalse(status_is_complete("not_found"))
        self.assertFalse(status_is_complete("msrp_not_found"))
        self.assertFalse(status_is_complete("error: TimeoutError"))

    def test_should_retry_crawl_only_for_non_not_found_statuses(self):
        self.assertFalse(should_retry_crawl("not_found", attempt_on_search=False))
        self.assertTrue(should_retry_crawl("not_found", attempt_on_search=True))
        self.assertFalse(should_retry_crawl("success", attempt_on_search=True))

    def test_resolve_column_supports_header_and_excel_letter(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(("Description", "OEM Number"))
        sheet.append(("Air filter", "12030-2N100"))
        self.assertEqual(resolve_column(sheet, "OEM Number", 1), (2, 2))
        self.assertEqual(resolve_column(sheet, "B", 1), (2, 2))
        workbook.close()

    def test_read_oem_values_from_workbook(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "parts.xlsx"
            workbook = Workbook()
            workbook.active.append(("Description", "OEM Number"))
            workbook.active.append(("Air filter", "12030-2N100"))
            workbook.save(path)
            workbook.close()
            self.assertEqual(read_oem_values(path, None, "OEM Number", 1), ["12030-2N100"])

    def test_duplicate_oems_are_queued_as_separate_rows(self):
        results, pending = prepare_crawl_rows(["12030-2N100", "12030-2N100"], [])
        self.assertEqual(len(results), 2)
        self.assertEqual(pending, [(0, "12030-2N100"), (1, "12030-2N100")])

    def test_save_and_load_results_round_trip(self):
        with TemporaryDirectory() as directory:
            output = Path(directory) / "results.xlsx"
            rows = [("12030-2N100", 30.06, date(2026, 10, 7), "success")]
            save_results(output, rows)
            workbook = load_workbook(output, data_only=True)
            self.assertEqual(
                tuple(cell.value for cell in workbook.active[1]),
                ("OEM Number", "MSRP", "Date", "Status"),
            )
            workbook.close()
            self.assertEqual(load_results(output)[0][0:2], ("12030-2N100", 30.06))


if __name__ == "__main__":
    unittest.main()
