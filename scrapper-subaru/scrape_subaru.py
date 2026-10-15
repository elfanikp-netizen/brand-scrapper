from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scrapper-shared"))

from parts_giant_crawler import CrawlerConfig, run


CONFIG = CrawlerConfig(
    brand="SubaruPartsDeal",
    site_url="https://www.subarupartsdeal.com/",
    cdp_port=9246,
    profile_name=".chrome-profile-subaru",
    part_label_pattern=r"(?i)Subaru\s+#?([A-Z0-9-]*\d[A-Z0-9-]*)",
)


if __name__ == "__main__":
    raise SystemExit(run(CONFIG, Path(__file__).resolve().parent, "Crawl SubaruPartsDeal MSRP values from an Excel OEM list."))