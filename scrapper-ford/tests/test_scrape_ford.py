import sys
import unittest
from pathlib import Path
from unittest.mock import Mock

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "scrapper-shared"))

from parts_giant_crawler import (
    CrawlerConfig,
    detail_status,
    detail_url_matches_oem,
    is_matching_detail_page,
    is_product_detail_page,
    parse_msrp,
)
from scrape_ford import CONFIG


class FordCrawlerTests(unittest.TestCase):
    def test_site_configuration(self):
        self.assertEqual(CONFIG.site_url, "https://www.fordpartsgiant.com/")
        self.assertEqual(CONFIG.cdp_port, 9245)

    def test_product_url_and_msrp(self):
        url = "https://www.fordpartsgiant.com/parts/ford-filter-asy-oil_fl-500-s.html"
        self.assertTrue(is_product_detail_page(url, CONFIG))
        self.assertTrue(detail_url_matches_oem(url, "FL-500S", CONFIG))
        self.assertFalse(detail_url_matches_oem(url, "FL-2055", CONFIG))
        self.assertEqual(parse_msrp("$8.14 MSRP: $12.71"), 12.71)

    def test_separate_site_configuration_type(self):
        self.assertIsInstance(CONFIG, CrawlerConfig)

    def test_replacement_detail_confirms_old_oem_without_consuming_next_label(self):
        page = Mock()
        page.url = "https://www.fordpartsgiant.com/parts/ford-filter-asy-oil_fl-500-s.html"
        page.locator.return_value.inner_text.return_value = (
            "Ford FL-500-S Oil Filter Assembly\n"
            "Replaces\nBU2Z-6731-A, FL-2055\nManufacturer\nFord\nMSRP: $12.71"
        )
        self.assertTrue(is_matching_detail_page(page, "FL-2055", CONFIG))
        self.assertEqual(
            detail_status("Ford FL-500-S Oil Filter Assembly", "FL-2055", 12.71, page.url, CONFIG),
            "success_superseded: fl-500-s",
        )


if __name__ == "__main__":
    unittest.main()