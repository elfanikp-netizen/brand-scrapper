# VW Parts MSRP Crawler

Reads OEM part numbers from an Excel workbook, searches [parts.vw.com](https://parts.vw.com/), opens a matching result, and writes the OEM number, displayed MSRP, and crawl date to a separate workbook.

## Setup

Use Python 3.10 or newer and install Google Chrome. From this folder, install the Python dependencies:

```powershell
python -m pip install -r requirements.txt
```

## Run

Select a column by its header, Excel letters, or 1-based number:

```powershell
python scrape_vw.py --input .\parts.xlsx --column "OEM Number"
python scrape_vw.py --input .\parts.xlsx --column B --output .\vw-prices.xlsx
```

By default, row 1 is treated as the header. For a sheet with no header, pass `--header-row 0`. Choose a worksheet with `--sheet "Sheet1"`. The default output is `<input>_results.xlsx` and contains `OEM Number`, `MSRP`, `Date`, and `Status` columns. Duplicate OEMs in the input are crawled for each row and saved as separate output rows. The workbook is saved after every part so completed rows remain available if the crawl is interrupted.

The crawler resumes by default when the output workbook already exists. It skips only successful input-row occurrences (`success` or `success_superseded: ...`); unsuccessful rows are retried on resume. Duplicate OEM occurrences are matched by their occurrence order, so a successful first duplicate does not cause later duplicate input rows to be skipped. Each row gets up to three attempts in a run. Use `--resume` to make resume behavior explicit, or `--start-over` to ignore existing results and crawl all input rows again:

```powershell
python scrape_vw.py --input .\parts.xlsx --column "OEM Number" --resume
python scrape_vw.py --input .\parts.xlsx --column "OEM Number" --start-over
```

Status values are `success`, `success_superseded: <current OEM>`, `not_found`, `msrp_not_found`, or `error: ...` with the exception details. VW may route a discontinued OEM to its replacement part; these rows retain the searched OEM in the first column, use the replacement's MSRP, and name the current OEM in Status. Older result workbooks without a `Status` column can be resumed; rows with an MSRP are treated as successful, and rows without one are retried.

The crawler opens the installed Google Chrome visibly and attaches Playwright through Chrome DevTools. It reuses one browser tab and VW's same `#SearchInput` header textbox for all OEM searches and product pages, then closes its dedicated Chrome session after the crawl finishes or stops. It stores the Chrome session under `.chrome-profile`, so verification cookies persist across runs. If VW presents a human-verification page, complete the check in the browser; the crawler detects when the page clears and continues automatically. It waits up to five minutes without reloading the challenge page. This tool does not bypass verification. Keep the profile private because it may contain session cookies. If Chrome is installed outside its standard location, pass `--chrome-path "C:\Program Files\Google\Chrome\Application\chrome.exe"`.

If Cloudflare returns to the verification screen after you complete it, check that your DNS, VPN, proxy, and firewall allow HTTPS access to `brunhild.challenges.cloudflare.com`. On Windows, test name resolution with `Resolve-DnsName brunhild.challenges.cloudflare.com -Type A`. The crawler reports a challenge-host network failure instead of repeatedly waiting when it detects one.

## Tests

```powershell
python -m unittest discover -s tests -v
```