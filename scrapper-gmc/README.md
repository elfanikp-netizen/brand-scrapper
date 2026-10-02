# GM Parts MSRP Crawler

Reads OEM/GM part numbers from an Excel workbook, searches [parts.gmparts.com](https://parts.gmparts.com/), opens the matching product card, and writes the queried OEM number, product MSRP, crawl date, and status to a separate workbook.

## Setup

Use Python 3.10 or newer and install Google Chrome. From this folder:

```powershell
python -m pip install -r requirements.txt
```

## Run

Select the input column by its header, Excel letters, or 1-based number:

```powershell
python scrape_gmc.py --input .\parts.xlsx --column "OEM Number"
python scrape_gmc.py --input .\parts.xlsx --column B --output .\gmc-prices.xlsx
```

Row 1 is treated as the header by default. Use `--header-row 0` for a headerless sheet or `--sheet "Sheet1"` to select a worksheet. The default output is `<input>_results.xlsx` with `OEM Number`, `MSRP`, `Date`, and `Status` columns. Duplicate OEM values in the input are crawled independently and appear as separate output rows.

The crawler resumes by default. It skips only successful input-row occurrences and retries `not_found`, `msrp_not_found`, and `error` statuses. Duplicate occurrences are matched to prior results by order; a completed result for one duplicate does not skip the others. Each input row gets up to three attempts, and each attempt is saved. Use `--resume` explicitly or `--start-over` to crawl all input rows again:

```powershell
python scrape_gmc.py --input .\parts.xlsx --column "OEM Number" --resume
python scrape_gmc.py --input .\parts.xlsx --column "OEM Number" --start-over
```

The visible installed Chrome session uses a dedicated `.chrome-profile`, a single tab, and a separate DevTools port. It closes after the run. If the site requests human verification, complete it in the browser; the crawler waits for the page to clear and does not bypass the check. Product suggestions can identify replaced part numbers; the searched OEM stays in the first output column and the current GM part number is recorded in a `success_superseded` status.

## Tests

```powershell
python -m unittest discover -s tests -v
```