import unittest
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from scrape_toyota import (
    load_results,
    crawl_msrp,
    detail_status,
    detail_url_matches_oem,
    dismiss_blocking_popups,
    is_toyota_site_url,
    normalize_part_number,
    parse_msrp,
    parse_product_msrp,
    prepare_crawl_rows,
    read_oem_values,
    resolve_column,
    save_results,
    should_retry_crawl,
    status_is_complete,
    wait_for_product_detail,
    wait_for_product_msrp,
    wait_for_product_suggestion,
)


class CrawlerHelpersTests(unittest.TestCase):
    def test_search_fallback_is_disabled_by_default(self):
        with patch("sys.argv", ["scrape_toyota.py", "--input", "parts.xlsx", "--column", "A"]):
            from scrape_toyota import parse_args

            args = parse_args()

        self.assertFalse(args.attempt_on_search)

    def test_search_fallback_flag_enables_full_search(self):
        with patch(
            "sys.argv",
            ["scrape_toyota.py", "--input", "parts.xlsx", "--column", "A", "--attempt-on-search"],
        ):
            from scrape_toyota import parse_args

            args = parse_args()

        self.assertTrue(args.attempt_on_search)

    def test_not_found_advances_to_next_oem_without_search_fallback(self):
        self.assertFalse(should_retry_crawl("not_found", attempt_on_search=False))

    def test_not_found_retries_when_search_fallback_is_enabled(self):
        self.assertTrue(should_retry_crawl("not_found", attempt_on_search=True))

    def test_success_never_retries(self):
        self.assertFalse(should_retry_crawl("success", attempt_on_search=True))

    def test_missing_product_suggestion_does_not_submit_search_by_default(self):
        steps = []

        class SearchBox:
            def fill(self, value):
                steps.append(f"fill:{value}")

            def press(self, key):
                steps.append(f"press:{key}")

        with (
            patch("scrape_toyota.dismiss_blocking_popups"),
            patch("scrape_toyota.find_search_box", return_value=SearchBox()),
            patch("scrape_toyota.wait_for_product_suggestion", return_value=None),
        ):
            msrp, status = crawl_msrp(object(), "8115002M90", [])

        self.assertEqual((msrp, status), (None, "not_found"))
        self.assertEqual(steps, ["fill:8115002M90"])

    def test_missing_product_suggestion_submits_search_when_enabled(self):
        steps = []

        class SearchBox:
            def fill(self, value):
                steps.append(f"fill:{value}")

            def press(self, key):
                steps.append(f"press:{key}")

        class Page:
            def wait_for_load_state(self, state, timeout):
                steps.append(f"load:{state}")

        with (
            patch("scrape_toyota.dismiss_blocking_popups"),
            patch("scrape_toyota.find_search_box", return_value=SearchBox()),
            patch("scrape_toyota.wait_for_product_suggestion", return_value=None),
            patch("scrape_toyota.wait_for_human_verification"),
            patch("scrape_toyota.is_matching_detail_page", return_value=False),
            patch("scrape_toyota.wait_for_matching_result", return_value=None),
        ):
            msrp, status = crawl_msrp(Page(), "8115002M90", [], attempt_on_search=True)

        self.assertEqual((msrp, status), (None, "not_found"))
        self.assertIn("press:Enter", steps)

    def test_crawl_clicks_product_suggestion_before_reading_detail_msrp(self):
        steps = []

        class SearchBox:
            def fill(self, value):
                steps.append(f"fill:{value}")

        class Suggestion:
            def click(self, timeout):
                steps.append("suggestion_click")

        def dismiss(page, press_escape=True):
            steps.append(f"dismiss:{press_escape}")

        product_text = "Toyota Genuine #81110-02M70\nRight Hand Headlamp\nMSRP $1,103.12"
        with (
            patch("scrape_toyota.dismiss_blocking_popups", side_effect=dismiss),
            patch("scrape_toyota.find_search_box", return_value=SearchBox()),
            patch(
                "scrape_toyota.wait_for_product_suggestion",
                side_effect=lambda page: steps.append("suggestion_wait") or Suggestion(),
            ),
            patch(
                "scrape_toyota.wait_for_human_verification",
                side_effect=lambda page, failures: steps.append("verification_wait"),
            ),
            patch(
                "scrape_toyota.wait_for_product_detail",
                side_effect=lambda page: steps.append("detail_wait") or True,
            ),
            patch(
                "scrape_toyota.wait_for_product_msrp",
                side_effect=lambda page: steps.append("pdp_msrp_wait") or product_text,
            ),
        ):
            msrp, status = crawl_msrp(object(), "8111002M70", [])

        self.assertEqual((msrp, status), (1103.12, "success"))
        self.assertEqual(
            steps,
            [
                "dismiss:True",
                "fill:8111002M70",
                "suggestion_wait",
                "suggestion_click",
                "verification_wait",
                "detail_wait",
                "dismiss:True",
                "pdp_msrp_wait",
            ],
        )

    def test_wait_for_product_detail_after_suggestion_click_navigation(self):
        class Product:
            def __init__(self, page):
                self.page = page

            def count(self):
                return 1 if self.page.waits >= 2 else 0

            def is_visible(self):
                return True

            def inner_text(self, timeout):
                return "Toyota Genuine #81110-02M70"

        class Page:
            waits = 0

            @property
            def url(self):
                if self.waits >= 2:
                    return "https://autoparts.toyota.com/products/product/headlamp-assy-rh-8111002m70"
                return "https://autoparts.toyota.com/"

            def locator(self, selector):
                self.selector = selector
                return Product(self)

            def wait_for_timeout(self, milliseconds):
                self.waits += 1

        page = Page()

        self.assertTrue(wait_for_product_detail(page, timeout_seconds=1))
        self.assertGreaterEqual(page.waits, 2)
        self.assertEqual(page.selector, 'main[aria-label="pdp container"]')

    def test_wait_for_product_msrp_until_pdp_content_is_rendered(self):
        class Product:
            def __init__(self, page):
                self.page = page

            def count(self):
                return 1

            def is_visible(self):
                return True

            def inner_text(self, timeout):
                if self.page.waits >= 2:
                    return "Toyota Genuine #81110-02M70\\nMSRP $1,103.12"
                return "Toyota Genuine #81110-02M70"

        class Page:
            waits = 0

            def locator(self, selector):
                self.selector = selector
                return Product(self)

            def wait_for_timeout(self, milliseconds):
                self.waits += 1

        page = Page()
        product_text = wait_for_product_msrp(page, timeout_seconds=1)

        self.assertIn("MSRP $1,103.12", product_text)
        self.assertGreaterEqual(page.waits, 2)
        self.assertEqual(page.selector, 'main[aria-label="pdp container"]')

    def test_popup_dismissal_does_not_send_escape_when_preserving_suggestion(self):
        class Button:
            def is_visible(self):
                return True

            def is_enabled(self):
                return True

            def click(self, timeout):
                self.page.dismissed += 1

            def __init__(self, page):
                self.page = page

        class Buttons:
            def __init__(self, page):
                self.page = page

            def count(self):
                return 1 if self.page.dismissed == 0 else 0

            def nth(self, index):
                return Button(self.page)

        class Keyboard:
            def __init__(self, page):
                self.page = page

            def press(self, key):
                self.page.keys.append(key)

        class Page:
            def __init__(self):
                self.dismissed = 0
                self.keys = []
                self.keyboard = Keyboard(self)

            def locator(self, selector):
                return Buttons(self)

        page = Page()
        dismiss_blocking_popups(page, press_escape=False)

        self.assertEqual(page.dismissed, 1)
        self.assertEqual(page.keys, [])

    def test_wait_for_product_suggestion_until_it_becomes_visible(self):
        class Suggestion:
            def is_visible(self):
                return True

            def get_attribute(self, name):
                return "https://autoparts.toyota.com/products/product/headlamp-assy-lh-8115002m90"

            def inner_text(self):
                return "Left Hand Headlamp Assembly"

        class Suggestions:
            def __init__(self, page):
                self.page = page

            def count(self):
                return 1 if self.page.waits >= 2 else 0

            def nth(self, index):
                return Suggestion()

        class Page:
            waits = 0

            def locator(self, selector):
                self.selector = selector
                return Suggestions(self)

            def wait_for_timeout(self, milliseconds):
                self.waits += 1

        page = Page()
        suggestion = wait_for_product_suggestion(page, timeout_seconds=1)

        self.assertIsNotNone(suggestion)
        self.assertGreaterEqual(page.waits, 2)
        self.assertIn('text-is("Product suggestions")', page.selector)

    def test_normalize_part_number_ignores_punctuation_and_case(self):
        self.assertEqual(normalize_part_number("90915-YZZF2"), "90915YZZF2")

    def test_detail_url_matches_queried_superseded_oem(self):
        self.assertTrue(
            detail_url_matches_oem(
                "https://autoparts.toyota.com/products/product/filter-s-a-oil-90915yzzn1",
                "90915-YZZN1",
            )
        )

    def test_detects_existing_toyota_page_to_avoid_duplicate_navigation(self):
        self.assertTrue(is_toyota_site_url("https://autoparts.toyota.com/"))
        self.assertTrue(
            is_toyota_site_url(
                "https://autoparts.toyota.com/products/product/filter-s-a-oil-90915yzzn1"
            )
        )
        self.assertFalse(is_toyota_site_url("about:blank"))
        self.assertFalse(
            detail_url_matches_oem(
                "https://autoparts.toyota.com/products/product/filter-s-a-oil-90915yzzn1",
                "90915-YZZF2",
            )
        )

    def test_superseded_detail_status_discloses_current_part_number(self):
        self.assertEqual(
            detail_status("Toyota Genuine #90915-YZZN1", "90915-YZZF2", 6.57),
            "success_superseded: 90915-YZZN1",
        )

    def test_category_label_is_not_mistaken_for_replacement_part(self):
        self.assertEqual(detail_status("Clearance\\nMSRP $1,104.45", "8115002M90", 1104.45), "success")

    def test_only_successful_statuses_are_complete(self):
        self.assertTrue(status_is_complete("success"))
        self.assertTrue(status_is_complete("success_superseded: 90915-YZZN1"))
        self.assertFalse(status_is_complete("success_superseded: Clearance"))
        self.assertFalse(status_is_complete("not_found"))
        self.assertFalse(status_is_complete("msrp_not_found"))
        self.assertFalse(status_is_complete("error: TimeoutError"))

    def test_duplicate_oems_are_crawled_as_separate_rows(self):
        results, pending = prepare_crawl_rows(
            ["1K0123456A", "1K0123456A"],
            [],
        )
        self.assertEqual(len(results), 2)
        self.assertEqual(pending, [(0, "1K0123456A"), (1, "1K0123456A")])

    def test_duplicate_resume_matches_successes_by_occurrence(self):
        saved = [
            ("1K0123456A", 10.0, date(2026, 10, 1), "success"),
            ("1K0123456A", None, date(2026, 10, 1), "not_found"),
        ]
        results, pending = prepare_crawl_rows(["1K0123456A", "1K0123456A"], saved)
        self.assertEqual(results[0][1:], saved[0][1:])
        self.assertEqual(pending, [(1, "1K0123456A")])

    def test_duplicate_rows_from_old_deduplicated_results_get_crawled(self):
        saved = [("1K0123456A", 10.0, date(2026, 10, 1), "success")]
        results, pending = prepare_crawl_rows(["1K0123456A", "1K0123456A"], saved)
        self.assertEqual(results[0][3], "success")
        self.assertEqual(pending, [(1, "1K0123456A")])

    def test_parse_msrp_reads_currency_and_thousands_separator(self):
        self.assertEqual(parse_msrp("Toyota Genuine #90915-YZZN1\nMSRP $6.57"), 6.57)

    def test_parse_msrp_requires_an_msrp_label(self):
        self.assertIsNone(parse_msrp("Price: $123.45"))

    def test_parse_product_msrp_ignores_unrelated_disclosure_prices(self):
        text = (
            "Toyota Genuine #81110-02M70\nRight Hand Headlamp\nMSRP $1,103.12\n"
            "Warranty details: battery MSRP $99.00"
        )
        self.assertEqual(parse_product_msrp(text), 1103.12)

    def test_resolve_column_supports_header_and_excel_letter(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(("Description", "OEM Number"))
        sheet.append(("Widget", "1K0123456A"))
        self.assertEqual(resolve_column(sheet, "OEM Number", 1), (2, 2))
        self.assertEqual(resolve_column(sheet, "B", 1), (2, 2))

    def test_read_oem_values_handles_headerless_sheet(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "parts.xlsx"
            workbook = Workbook()
            workbook.active.append(("1K0123456A",))
            workbook.active.append(("5Q0123456B",))
            workbook.save(path)
            self.assertEqual(read_oem_values(path, None, "A", 0), ["1K0123456A", "5Q0123456B"])

    def test_save_results_creates_status_column_and_round_trips_rows(self):
        with TemporaryDirectory() as directory:
            output = Path(directory) / "results.xlsx"
            results = [
                ("1K0123456A", 12.5, date(2026, 10, 1), "success"),
                ("5Q0123456B", None, date(2026, 10, 1), "error: TimeoutError"),
            ]
            save_results(output, results)
            workbook = load_workbook(output, data_only=True)
            self.assertEqual(
                tuple(cell.value for cell in workbook.active[1]),
                ("OEM Number", "MSRP", "Date", "Status"),
            )
            row = tuple(cell.value for cell in workbook.active[2])
            self.assertEqual(row[:2], ("1K0123456A", 12.5))
            self.assertEqual(row[2].date(), date(2026, 10, 1))
            self.assertEqual(row[3], "success")
            workbook.close()

            loaded = load_results(output)
            self.assertEqual(loaded[0][0:2], ("1K0123456A", 12.5))
            self.assertEqual(loaded[0][3], "success")
            self.assertEqual(loaded[1][3], "error: TimeoutError")
            self.assertTrue(status_is_complete(loaded[0][3]))
            self.assertFalse(status_is_complete(loaded[1][3]))

    def test_legacy_rows_without_status_retry_when_msrp_is_empty(self):
        with TemporaryDirectory() as directory:
            output = Path(directory) / "legacy.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(("OEM Number", "MSRP", "Date"))
            sheet.append(("90915-YZZF2", None, date(2026, 10, 1)))
            sheet.append(("85214-02340", 10.81, date(2026, 10, 1)))
            workbook.save(output)
            loaded = load_results(output)
            self.assertEqual(loaded[0][3], "error: legacy row has no status or MSRP")
            self.assertEqual(loaded[1][3], "success")


if __name__ == "__main__":
    unittest.main()