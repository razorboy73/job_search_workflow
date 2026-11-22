# Workflow: Scrape DailyRemote job search results into a spreadsheet

## Objective
Collect every job from a DailyRemote search (e.g. "social media") into one Excel workbook with all useful fields, then put a copy in Google Sheets so the user can filter, sort and track applications.

## Required inputs
| Input | Example | Notes |
|---|---|---|
| Search term | `social media` | The `?search=` value in the DailyRemote URL |
| Number of listing pages | `23` | Shown at the bottom of the search page. Each page has 30 jobs. If unsure, overshoot: the tool stops at the first empty page. |
| Depth | listings + job pages | See "Cost" below. Ask the user if the credit cost is large. |

The user may give a URL like `https://dailyremote.com/?search=social+media&page=1`. Take the search term from it.

## Tool
`tools/scrape_dailyremote.py`

```
.venv\Scripts\python.exe tools\scrape_dailyremote.py --search "social media" --pages 23
```

| Flag | Purpose |
|---|---|
| `--pages N` | Listing pages to scrape |
| `--skip-details` | Listing pages only (cheap, but no company names) |
| `--max-details N` | Only open the first N job pages (use for dry runs) |
| `--limit N` | Keep only the first N unique jobs in the sheet, all with job page details. Listing pages are still scraped so the Run Log can report the full count. |
| `--workers 3` | Parallel Firecrawl requests |
| `--out path.xlsx` | Default: `dailyremote_<search>_jobs_<date>.xlsx` in the project root |

The tool needs `FIRECRAWL_API_KEY` in `.env`.

## Steps
1. **Check credits.** Use the Firecrawl MCP `firecrawl_credit_usage` tool. A full run costs about one credit per listing page plus one per job.
2. **Confirm with the user** whether to open job pages (cost and fields below), and where the output should go.
3. **Dry run:** `--pages 1 --max-details 3` (about 4 credits). Open the output and check by eye that location, salary, category, role, company and description are filled and correctly split.
4. **Full run.** Run in the background because it takes 15 to 30 minutes for about 650 jobs. Progress goes to stdout; redirect it to `.tmp/dailyremote/run.log`.
5. **Verify** using the workbook's **Run Log** sheet:
   - "Unique jobs in sheet" should be close to "Jobs reported by site". A gap of a few is normal, because listings change during the run.
   - "Job pages failed" should be 0 or near 0. If not, rerun the same command: cached pages are free, so only the failed ones cost credits.
   - Spot-check 2 or 3 rows against the live site.
6. **Upload** to Google Drive using the Drive connector's `create_file`. Upload the .xlsx (`contentMimeType` `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`) with conversion on, so it becomes a Google Sheet. Give the user the link.

## Outputs
- `dailyremote_<search>_jobs_<date>.xlsx`, with two sheets:
  - **Jobs:** Job ID, Title, Company, Location, Job Type, Experience, Salary (as listed), Salary Min/Max/Period/Currency, Category, Role, Posted, Posted (est. date), AI Summary, Tags, LinkedIn Note, Job URL, Apply URL, Company URL, Full Description, Source Page, plus tracking columns Status (dropdown), Applied Date, Notes.
  - **Run Log:** counts, errors and credits used.
- A Google Sheet copy in Drive.
- Intermediates in `.tmp/dailyremote/`: `raw/` page cache, `<search>_jobs.csv` and `run.log`. All disposable.

## Cost
- Listing pages only: about 1 credit per page (23 pages is about 23 credits).
- Job pages: about 1 credit per job (650 jobs is about 650 credits). They add company name, company URL, apply link and full description.
- Reruns are nearly free: every fetched page is cached in `.tmp/dailyremote/raw/`. Delete that folder for a fresh scrape, for example on a new day.

## Site quirks learned (Nov 2026)
- **Listing pages hide the company** ("Company hidden"; real names are a Premium feature). Job pages show it in a `/remote-company/` link, and the page title "X at Company" is a fallback.
- **Location, salary and experience are glued together** in one line, e.g. `United States$93000 - $122K per year5-10 yrs exp`. The salary regex only lets a currency symbol or 3-letter code come before the number. An earlier version ate the end of the location ("United Stat" + "es$93000").
- **"Virtual Assistant" is both a category and a role.** The first link after the location is always the category; any later links are roles.
- Salaries come in `per year`, `per month` and `per hour`, with or without `K`. Filter on Salary Period before comparing Min/Max.
- "Posted" is relative ("3 Weeks Ago"), so Posted (est. date) is approximate. The job page's "Posted N days ago" is used when available because it is fresher.
- The last page has fewer than 30 jobs, and the site's total can be off by one or two versus what's parsed.
- Apply URLs carry a `nonce` and may expire. The Job URL is the stable link.
- **Some job pages have no "AI Summary" or "Description" heading** (about 11% in the Nov 2026 run). The posting text starts right after the resume-match widget's closing `×` line. The parser falls back to that position.

## Firecrawl limits learned (Nov 2026)
- **The plan allows about 20 requests a minute.** Above that, the API returns HTTP 429 with "retry after Ns". The tool waits for that time plus a few seconds, with up to 10 attempts. Real throughput is about 13 job pages a minute with 3 workers, so 650 jobs takes about 50 minutes. Raising `--workers` won't help.
- For a quicker result, use `--limit 100` (about 8 minutes) and grow it later. Cached pages make the bigger rerun cost only the new pages.
