from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scrapper-shared"))

from parts_giant_crawler import CrawlerConfig, run


CONFIG = CrawlerConfig(
    brand="FordPartsGiant",
    site_url="https://www.fordpartsgiant.com/",
    cdp_port=9245,
    profile_name=".chrome-profile",
    part_label_pattern=r"(?i)Ford\s+#?([A-Z0-9-]*\d[A-Z0-9-]*)",
)


if __name__ == "__main__":
    raise SystemExit(run(CONFIG, Path(__file__).resolve().parent, "Crawl FordPartsGiant MSRP values from an Excel OEM list."))