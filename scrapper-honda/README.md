# Honda Parts MSRP Crawler

Reads OEM Honda part numbers from an Excel workbook, searches [Honda OEM Parts Online](https://honda.oempartsonline.com/), opens the matching product page, and writes the queried OEM number, product MSRP, crawl date, and status to a separate workbook.

## Setup

Use Python 3.10 or newer and install Google Chrome. From this folder:

```powershell
python -m pip install -r requirements.txt
```

## Run

Select the input column by its header, Excel letters, or 1-based number:

```powershell
python scrape_honda.py --input .\parts.xlsx --column "OEM Number"
python scrape_honda.py --input .\parts.xlsx --column B --output .\honda-prices.xlsx
```

Row 1 is treated as the header by default. Use `--header-row 0` for a headerless sheet or `--sheet "Sheet1"` to select a worksheet. The default output is `<input>_results.xlsx` with `OEM Number`, `MSRP`, `Date`, and `Status` columns. Duplicate OEM values in the input are crawled independently and appear as separate output rows.

The crawler resumes by default. It skips only successful input-row occurrences. Each pending OEM is searched once by default; unsuccessful rows are saved and left for a later run. Use `--attempt-on-search` to retry unsuccessful searches up to three times in the current run. Duplicate occurrences are matched to prior results by order; a completed result for one duplicate does not skip the others. Use `--resume` explicitly or `--start-over` to crawl all input rows again:

```powershell
python scrape_honda.py --input .\parts.xlsx --column "OEM Number" --resume
python scrape_honda.py --input .\parts.xlsx --column "OEM Number" --start-over
python scrape_honda.py --input .\parts.xlsx --column "OEM Number" --attempt-on-search
```

The crawler opens a visible Chrome session with a dedicated `.chrome-profile`, searches each OEM through the site's search box, and submits the search to open its matching product page. It closes Chrome after the run. If the site requests human verification, complete it in the browser; the crawler waits for the page to clear and does not bypass the check. The searched OEM stays in the first output column; when the product page identifies a different manufacturer part number, it is recorded in a `success_superseded` status.

## Tests

```powershell
python -m unittest discover -s tests -v
```