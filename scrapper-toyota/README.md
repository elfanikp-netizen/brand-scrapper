# Toyota Parts MSRP Crawler

Reads OEM part numbers from an Excel workbook, searches [autoparts.toyota.com](https://autoparts.toyota.com/), opens the top ranked product result, and records the displayed MSRP in a separate workbook. When Toyota returns a replacement part, the searched OEM remains in the first column and the replacement number is recorded in Status.

## Setup

Use Python 3.10 or newer and install Google Chrome. From this folder, install the dependencies:

```powershell
python -m pip install -r requirements.txt
```

## Run

Select the OEM column by its header, Excel letters, or 1-based number:

```powershell
python scrape_toyota.py --input .\parts.xlsx --column "OEM Number"
python scrape_toyota.py --input .\parts.xlsx --column B --output .\toyota-prices.xlsx
```

Row 1 is treated as the header by default. Use `--header-row 0` for a headerless sheet or `--sheet "Sheet1"` to select a worksheet. The default output is `<input>_results.xlsx` with `OEM Number`, `MSRP`, `Date`, and `Status` columns. Duplicate OEMs are retained as separate rows, and the workbook is saved after each part.

The crawler resumes from an existing output by default. Successful rows (`success` or `success_superseded: ...`) are skipped; unsuccessful rows are retried up to three times. Use `--start-over` to recrawl every input row:

```powershell
python scrape_toyota.py --input .\parts.xlsx --column "OEM Number" --resume
python scrape_toyota.py --input .\parts.xlsx --column "OEM Number" --start-over
```

By default, the crawler only clicks a product from the autocomplete **Product suggestions** section. If no such suggestion appears, it records `not_found` and moves to the next OEM without submitting a full search or retrying that OEM in the current run. Use `--attempt-on-search` to opt into pressing Enter and checking the search-results page as a fallback:

```powershell
python scrape_toyota.py --input .\parts.xlsx --column "OEM Number" --attempt-on-search
```

Statuses are `success`, `success_superseded: <current OEM>`, `not_found`, `msrp_not_found`, or `error: ...`. Older result workbooks without a Status column can be resumed; rows with an MSRP are treated as successful.

The crawler opens Google Chrome visibly, attaches through Chrome DevTools, and uses Toyota's header search for each OEM. It keeps a dedicated profile in `.chrome-profile` so browser verification state persists between runs. If the site requests human verification, complete it in the browser; the crawler waits up to five minutes and does not attempt to bypass the check. Keep the profile private because it may contain session cookies. Pass `--chrome-path "C:\Program Files\Google\Chrome\Application\chrome.exe"` if Chrome is installed elsewhere.

## Tests

```powershell
python -m unittest discover -s tests -v
```