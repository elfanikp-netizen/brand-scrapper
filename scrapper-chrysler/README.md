# Chrysler / Mopar Parts MSRP Crawler

Reads OEM part numbers from an Excel workbook, searches [mymoparparts.com](https://www.mymoparparts.com/), opens a matching or closely related product result, and writes the searched OEM number, MSRP, crawl date, and status to a separate workbook.

## Setup

Use Python 3.10 or newer and install Google Chrome. From this folder, install the Python dependencies:

```powershell
python -m pip install -r requirements.txt
```

## Run

Choose the input column by its header, Excel letters, or 1-based number:

```powershell
python scrape_chrysler.py --input .\parts.xlsx --column "OEM Number"
python scrape_chrysler.py --input .\parts.xlsx --column B --output .\mopar-prices.xlsx
```

For MoparPartsGiant, use the alternate crawler:

```powershell
python scrape_chrysler2.py --input .\parts.xlsx --column "OEM Number"
```

Row 1 is treated as the header by default. Use `--header-row 0` for a headerless worksheet and `--sheet "Sheet1"` to select a worksheet. The default output is `<input>_results.xlsx`, with `OEM Number`, `MSRP`, `Date`, and `Status` columns. The workbook is saved after each part. Resume is enabled by default and skips successful rows; use `--start-over` to crawl every input row again.

The crawler submits the part number through Mopar's header search, prefers an exact result, and otherwise ranks related product OEM numbers by edit distance with the site's result order breaking ties. It opens the selected `/oem-parts/` page and reads MSRP from its detail text. Status values are `success`, `success_superseded: <current OEM>`, `not_found`, `msrp_not_found`, or `error: ...`.

The crawler opens Chrome visibly and attaches through Chrome DevTools. If the site asks for human verification, complete it in the browser; the crawler waits up to five minutes and leaves the browser open if verification does not clear. It does not bypass verification. The dedicated Chrome profile is stored in `.chrome-profile`. If Chrome is installed elsewhere, pass `--chrome-path` with its executable path.

## Tests

```powershell
python -m unittest discover -s tests -v
```