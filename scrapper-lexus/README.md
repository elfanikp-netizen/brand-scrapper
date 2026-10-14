# Lexus Parts MSRP Crawler

Reads Lexus OEM part numbers from Excel, searches [LexusPartsNow](https://www.lexuspartsnow.com/), opens the matching part page, and records MSRP, crawl date, and status.

## Setup

Use Python 3.10 or newer, install Google Chrome, and install dependencies from this folder:

```powershell
python -m pip install -r requirements.txt
```

## Run

```powershell
python scrape_lexus.py --input .\parts.xlsx --column "OEM Number"
python scrape_lexus.py --input .\parts.xlsx --column B --output .\lexus-prices.xlsx
```

The crawler writes `OEM Number`, `MSRP`, `Date`, and `Status`. It resumes successful rows by default; pass `--start-over` to recrawl all rows. Use `--sheet`, `--header-row 0`, or `--chrome-path` when needed. Results are saved after each part.

The search submits through the site's header and uses direct product navigation or a matching `/parts/` link. It verifies the product number (or a listed replacement) and reads MSRP from the detail page. Human verification must be completed in the visible browser; the crawler waits up to five minutes and does not bypass it.

Run tests from this folder with `python -m unittest discover -s tests -v`.