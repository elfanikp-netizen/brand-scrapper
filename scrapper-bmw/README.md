# BMW Parts MSRP Crawler

Reads BMW part numbers from an Excel workbook, searches [getbmwparts.com](https://www.getbmwparts.com/), opens the exact matching search result, and writes the searched OEM number, MSRP, crawl date, and status to a separate workbook.

## Setup

Use Python 3.10 or newer and install Google Chrome. From this folder, install the dependencies:

```powershell
python -m pip install -r requirements.txt
```

## Run

Choose the input column by its header, Excel letters, or 1-based number:

```powershell
python scrape_bmw.py --input .\parts.xlsx --column "OEM Number"
python scrape_bmw.py --input .\parts.xlsx --column B --output .\bmw-prices.xlsx
```

Row 1 is treated as the header by default. Use `--header-row 0` for a headerless worksheet and `--sheet "Sheet1"` to select a worksheet. The default output is `<input>_results.xlsx`, with `OEM Number`, `MSRP`, `Date`, and `Status` columns. The workbook is saved after each part. Resume is enabled by default and skips successful rows; use `--start-over` to crawl every input row again.

The crawler searches each part number and prefers a result whose displayed part number or product URL matches the input exactly. When an exact search returns multiple nearby OEMs, it ranks product numbers by edit distance to the query and uses the site's result order to break ties. A lone replacement card is accepted only when the results URL confirms the exact submitted OEM; with multiple cards, a candidate must be sufficiently close to the searched number. The selected card's detail URL confirms navigation, including when BMW does not list a `Replaces` field. MSRP is read from the detail page. Status values are `success`, `success_superseded: <current OEM>`, `not_found`, `msrp_not_found`, or `error: ...`.

The crawler opens Chrome visibly and attaches through Chrome DevTools. If the site asks for human verification, complete it in the browser; the crawler waits up to five minutes and leaves the browser open if verification does not clear. It does not bypass verification. The dedicated Chrome profile is stored in `.chrome-profile`. If Chrome is installed elsewhere, pass `--chrome-path` with its executable path.

## Tests

```powershell
python -m unittest discover -s tests -v
```