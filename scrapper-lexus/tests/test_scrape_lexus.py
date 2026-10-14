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
from scrape_lexus import CONFIG


class LexusCrawlerTests(unittest.TestCase):
    def test_site_configuration(self):
        self.assertEqual(CONFIG.site_url, "https://www.lexuspartsnow.com/")
        self.assertEqual(CONFIG.cdp_port, 9244)

    def test_product_url_and_msrp(self):
        url = "https://www.lexuspartsnow.com/parts/lexus-replaceable-element~04152-yzza1.html"
        self.assertTrue(is_product_detail_page(url, CONFIG))
        self.assertTrue(detail_url_matches_oem(url, "04152-YZZA1", CONFIG))
        self.assertFalse(detail_url_matches_oem(url, "04152-31090", CONFIG))
        self.assertEqual(parse_msrp("$5.07 MSRP: $6.57"), 6.57)

    def test_separate_site_configuration_type(self):
        self.assertIsInstance(CONFIG, CrawlerConfig)

    def test_replacement_detail_confirms_old_oem_without_consuming_next_label(self):
        page = Mock()
        page.url = "https://www.lexuspartsnow.com/parts/lexus-filter~04152-yzza1.html"
        page.locator.return_value.inner_text.return_value = (
            "Lexus 04152-YZZA1 Filter Element\n"
            "Replaces\n04152-31090, 04152-31110\nManufacturer\nLexus\nMSRP: $6.57"
        )
        self.assertTrue(is_matching_detail_page(page, "04152-31090", CONFIG))
        self.assertEqual(
            detail_status("Lexus 04152-YZZA1 Filter Element", "04152-31090", 6.57, page.url, CONFIG),
            "success_superseded: 04152-yzza1",
        )


if __name__ == "__main__":
    unittest.main()