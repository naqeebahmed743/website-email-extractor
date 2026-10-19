# Website Email Extractor

Multithreaded Python tool that takes an Excel file of website URLs, visits each site,
finds publicly listed contact email addresses and writes them back to Excel.

I first built it during my freelance work on B2B lead-list projects (processing lists of
1,000+ websites per run), then refactored it into a reusable command-line tool.

## How it works

1. Reads column A (website) and optional column B (contact page URL) from every sheet.
2. Fetches pages in parallel with a thread pool (`concurrent.futures`).
3. Extracts emails from the page text and from `mailto:` links, removes duplicates and
   false matches such as image names (`logo@2x.png`).
4. If the homepage has no email, falls back to the contact page.
5. Writes `emails` and `status` columns to a new Excel file and lists unreachable
   sites in `failed_urls.txt`.

All Excel writes happen in the main thread, so the workbook is never modified by
several threads at once.

## Example

Input (`sample/websites.xlsx`):

| website | contact_page |
|---|---|
| https://example-bakery.de | https://example-bakery.de/contact |
| https://example-studio.com | https://example-studio.com/contact |

Output (`sample/websites_results.xlsx`):

| website | contact_page | emails | status |
|---|---|---|---|
| https://example-bakery.de | … | info@example-bakery.de, orders@example-bakery.de | found (main page) |
| https://example-studio.com | … | hello@example-studio.com | found (contact page) |

## How to run

```bash
pip install -r requirements.txt
python email_extractor.py sample/websites.xlsx
python email_extractor.py websites.xlsx --output results.xlsx --workers 10 --timeout 15
```

## Tech

Python, Requests, BeautifulSoup, openpyxl, `concurrent.futures.ThreadPoolExecutor`, regex

## Responsible use

Only collect business contact addresses that are publicly listed, respect each site's
`robots.txt` and terms, keep request rates low, and follow GDPR / local data-protection
rules when storing or using the results.
