# Audi Parts MSRP Crawler

Reads Audi OEM part numbers from Excel, searches [AudiPartsGiant](https://www.audipartsgiant.com/), opens the matching genuine part page, and records its MSRP and crawl date. When a replacement part is returned, the requested OEM stays in the first column and the current manufacturer part number is recorded in Status.

## Setup

Use Python 3.10 or newer and install Google Chrome. From this folder, install the dependencies:

```powershell
python -m pip install -r requirements.txt
```

## Run

Select the OEM column by its header, Excel letters, or 1-based number:

```powershell
python scrape_audi.py --input .\parts.xlsx --column "OEM Number"
python scrape_audi.py --input .\parts.xlsx --column B --output .\audi-prices.xlsx
```

Row 1 is treated as the header by default. Use `--header-row 0` for a headerless worksheet or `--sheet "Sheet1"` to select a worksheet. The default output is `<input>_results.xlsx` with `OEM Number`, `MSRP`, `Date`, and `Status` columns. Duplicate OEMs are kept as separate rows, and results are saved after each part.

The crawler resumes from an existing output by default and skips successful rows (`success` or `success_superseded: ...`) that have an MSRP. Each OEM is searched once per run. Use `--start-over` to recrawl every input row:

```powershell
python scrape_audi.py --input .\parts.xlsx --column "OEM Number" --resume
python scrape_audi.py --input .\parts.xlsx --column "OEM Number" --start-over
```

The crawler closes available popups, submits each part number through AudiPartsGiant's header search, and accepts direct `/genuine/` product pages or matching product links from search results. MSRP is read from the product price area, including unavailable products.

Statuses are `success`, `success_superseded: <current OEM>`, `not_found`, `msrp_not_found`, or `error: ...`. Older result workbooks without a Status column can be resumed; rows with an MSRP are treated as successful.

The crawler opens Google Chrome visibly and attaches through Chrome DevTools. It uses a dedicated `.chrome-profile` so browser verification state persists between runs. Complete any human verification in the visible browser; the crawler waits for the page to clear and does not bypass the check. Keep the profile private because it may contain session cookies. Pass `--chrome-path "C:\Program Files\Google\Chrome\Application\chrome.exe"` if Chrome is installed elsewhere.

## Tests

```powershell
python -m unittest discover -s tests -v
```
