# Hyundai Parts MSRP Crawler

Reads Hyundai OEM part numbers from an Excel workbook, searches [HyundaiPartsDeal](https://www.hyundaipartsdeal.com/), opens the matching product page, and records the discounted selling price in the `MSRP` column (falling back to MSRP when there is no discount). When a replacement part is returned, the searched OEM remains in the first column and the current part number is recorded in Status.

## Setup

Use Python 3.10 or newer and install Google Chrome. From this folder, install the dependencies:

```powershell
python -m pip install -r requirements.txt
```

## Run

Select the OEM column by its header, Excel letters, or 1-based number:

```powershell
python scrape_hyundai.py --input .\parts.xlsx --column "OEM Number"
python scrape_hyundai.py --input .\parts.xlsx --column B --output .\hyundai-prices.xlsx
```

Row 1 is treated as the header by default. Use `--header-row 0` for a headerless sheet or `--sheet "Sheet1"` to select a worksheet. The default output is `<input>_results.xlsx` with `OEM Number`, `MSRP`, `Date`, and `Status` columns. Duplicate OEMs are retained as separate rows, and the workbook is saved after each part.

The crawler resumes from an existing output by default. Successful rows (`success` or `success_superseded: ...`) are skipped. It submits each OEM once; if the search does not reach a matching detail page, the crawler closes any popup and continues with the next value. Use `--start-over` to recrawl every input row:

```powershell
python scrape_hyundai.py --input .\parts.xlsx --column "OEM Number" --resume
python scrape_hyundai.py --input .\parts.xlsx --column "OEM Number" --start-over
```

HyundaiPartsDeal's header search opens the product detail directly for an exact OEM match. If full search opens a results page, the crawler selects the result that matches the requested OEM.

Statuses are `success`, `success_superseded: <current OEM>`, `not_found`, `msrp_not_found`, or `error: ...`. Older result workbooks without a Status column can be resumed; rows with an MSRP are treated as successful.

The crawler opens Google Chrome visibly, attaches through Chrome DevTools, and searches each OEM through HyundaiPartsDeal's header search. It keeps a dedicated profile in `.chrome-profile` so browser verification state persists between runs. If the site requests human verification, complete it in the browser; the crawler waits for the page to clear and does not attempt to bypass the check. Keep the profile private because it may contain session cookies. Pass `--chrome-path "C:\Program Files\Google\Chrome\Application\chrome.exe"` if Chrome is installed elsewhere.

## Tests

```powershell
python -m unittest discover -s tests -v
```