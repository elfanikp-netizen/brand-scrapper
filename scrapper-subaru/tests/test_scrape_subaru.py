import sys
import unittest
from pathlib import Path
from unittest.mock import Mock

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "scrapper-shared"))

from parts_giant_crawler import (
    CrawlerConfig,
    detail_url_matches_oem,
    is_matching_detail_page,
    is_product_detail_page,
    parse_msrp,
)
from scrape_subaru import CONFIG


class SubaruCrawlerTests(unittest.TestCase):
    def test_site_configuration(self):
        self.assertEqual(CONFIG.site_url, "https://www.subarupartsdeal.com/")
        self.assertEqual(CONFIG.cdp_port, 9246)
        self.assertEqual(CONFIG.profile_name, ".chrome-profile-subaru")

    def test_product_url_and_msrp(self):
        url = "https://www.subarupartsdeal.com/parts/subaru-filter-oil~15208aa15a.html"
        self.assertTrue(is_product_detail_page(url, CONFIG))
        self.assertTrue(detail_url_matches_oem(url, "15208AA15A", CONFIG))
        self.assertFalse(detail_url_matches_oem(url, "15208AA21A", CONFIG))
        self.assertEqual(parse_msrp("$7.02 MSRP: $9.75 You Save: $2.73"), 9.75)

    def test_configuration_type(self):
        self.assertIsInstance(CONFIG, CrawlerConfig)

    def test_matching_detail_uses_subaru_part_label(self):
        page = Mock()
        page.url = "https://www.subarupartsdeal.com/parts/subaru-filter-oil~15208aa15a.html"
        page.locator.return_value.inner_text.return_value = (
            "Subaru 15208AA15A Oil Filter Complete\n"
            "Manufacturer Part Number 15208AA15A\n"
            "$7.02 MSRP: $9.75"
        )
        self.assertTrue(is_matching_detail_page(page, "15208AA15A", CONFIG))
        self.assertFalse(is_matching_detail_page(page, "15208AA21A", CONFIG))


if __name__ == "__main__":
    unittest.main()