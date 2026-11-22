r"""
Scrape DailyRemote search results (listing pages + each job's detail page) into an Excel workbook.

- Fetches pages through the Firecrawl REST API (FIRECRAWL_API_KEY in .env).
- Caches every fetched page in .tmp/dailyremote/raw/ so reruns don't spend credits twice.
- Writes a "Jobs" sheet (with salary parsing + status-tracking columns) and a "Run Log" sheet,
  plus a CSV copy in .tmp/dailyremote/jobs.csv.

Run with:
  .venv\Scripts\python.exe tools\scrape_dailyremote.py --search "social media" --pages 23
  .venv\Scripts\python.exe tools\scrape_dailyremote.py --pages 1 --max-details 3   # cheap dry run
"""

import argparse
import csv
import json
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import quote_plus

import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
TMP = ROOT / ".tmp"
load_dotenv(ROOT / ".env")

BASE = "https://dailyremote.com"
API = "https://api.firecrawl.dev/v2/scrape"
WORK_DIR = TMP / "dailyremote"
RAW = WORK_DIR / "raw"

CATEGORIES = {
    "Software Development", "Design", "Support", "Sales", "Writing", "Product", "Legal",
    "Finance", "Marketing", "Data Entry", "Healthcare", "Recruitment", "Teaching",
    "Virtual Assistant", "Others",
}
CURRENCIES = {"$": "USD", "€": "EUR", "£": "GBP", "₹": "INR", "¥": "JPY", "C$": "CAD", "A$": "AUD"}
STATUS_OPTIONS = ["Not applied", "Applied", "Interviewing", "Offer", "Rejected"]

credits_used = 0


# ---------------------------------------------------------------- fetching

def fetch(url: str, cache_file: Path, api_key: str) -> dict:
    """Return {'markdown', 'title'} for url, using the cache when present."""
    global credits_used
    if cache_file.exists():
        return json.loads(cache_file.read_text(encoding="utf-8"))
    body = {"url": url, "formats": ["markdown"], "onlyMainContent": True}
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    last_err = None
    for attempt in range(10):
        try:
            r = requests.post(API, json=body, headers=headers, timeout=90)
            if r.status_code in (429, 500, 502, 503, 504):
                last_err = f"HTTP {r.status_code}: {r.text[:200]}"
                # Firecrawl's 429 says "retry after Ns" (plan limit ~20 req/min); honor it plus jitter.
                m = re.search(r"retry after (\d+)s", r.text)
                wait = int(m.group(1)) + 3 if m else 0
                time.sleep(max(wait, 6 * (attempt + 1)))
                continue
            if r.status_code == 402:
                raise RuntimeError(f"Firecrawl says credits are exhausted: {r.text[:200]}")
            r.raise_for_status()
            data = r.json().get("data", {})
            meta = data.get("metadata", {}) or {}
            credits_used += int(meta.get("creditsUsed", 1) or 0)
            out = {"markdown": data.get("markdown", ""), "title": meta.get("title") or meta.get("ogTitle") or ""}
            if not out["markdown"]:
                raise RuntimeError("empty markdown")
            cache_file.write_text(json.dumps(out), encoding="utf-8")
            return out
        except requests.RequestException as e:
            last_err = str(e)
            time.sleep(5 * (attempt + 1))
    raise RuntimeError(f"failed after retries: {last_err}")


# ---------------------------------------------------------------- parsing helpers

LINK_RE = re.compile(r"\[([^\]]*)\]\(([^)]+)\)")


def strip_md(text: str) -> str:
    text = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", text)          # images
    text = LINK_RE.sub(r"\1", text)                            # links -> text
    text = text.replace("**", "").replace("\\", "")
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def parse_money(s: str):
    s = s.strip().replace(",", "")
    m = re.search(r"(\d+(?:\.\d+)?)\s*([KkMm]?)", s)
    if not m:
        return None
    v = float(m.group(1))
    if m.group(2).lower() == "k":
        v *= 1_000
    elif m.group(2).lower() == "m":
        v *= 1_000_000
    return int(v) if v == int(v) else v


def parse_salary(text: str) -> dict:
    out = {"salary_min": None, "salary_max": None, "salary_period": "", "currency": ""}
    if not text:
        return out
    pm = re.search(r"per\s+(\w+)", text)
    out["salary_period"] = pm.group(1).lower() if pm else ""
    cm = re.match(r"\s*([^\d\s]+)", text)
    if cm:
        sym = cm.group(1).strip()
        out["currency"] = CURRENCIES.get(sym, sym)
    amount = text[: pm.start()] if pm else text
    parts = [p for p in re.split(r"\s+-\s+|\s*–\s*", amount) if re.search(r"\d", p)]
    if parts:
        out["salary_min"] = parse_money(parts[0])
        out["salary_max"] = parse_money(parts[-1])
    return out


def estimate_date(age: str, today: date):
    a = age.lower().strip()
    if not a:
        return None
    if "today" in a or "hour" in a or "minute" in a or "just" in a:
        return today
    if "yesterday" in a:
        return today - timedelta(days=1)
    m = re.search(r"(\d+)\s*(day|week|month|year)", a)
    if not m:
        return None
    n, unit = int(m.group(1)), m.group(2)
    days = {"day": 1, "week": 7, "month": 30, "year": 365}[unit] * n
    return today - timedelta(days=days)


# ---------------------------------------------------------------- listing pages

JOB_HEAD_RE = re.compile(r"^## \[(.+?)\]\((https://dailyremote\.com/remote-job/[^)\s]+)\)\s*$", re.M)
EXP_RE = re.compile(r"(\d+\s*-\s*\d+|\d+\+?)\s*yrs?\s*exp", re.I)
# Currency prefix is a symbol ($, €, ...) or a 3-letter code; it must not eat the tail of the location
# text it is glued to (e.g. "United States$93000 - $122K per year").
_CUR = r"(?:[$€£₹¥]|\b[A-Z]{3}\s?)"
SAL_RE = re.compile(rf"{_CUR}?\d[\d,.]*[KkMm]?(?:\s*-\s*{_CUR}?\d[\d,.]*[KkMm]?)?\s*per\s+\w+")


def parse_listing(md: str, page_no: int) -> list[dict]:
    jobs = []
    heads = list(JOB_HEAD_RE.finditer(md))
    for i, h in enumerate(heads):
        end = heads[i + 1].start() if i + 1 < len(heads) else len(md)
        block = md[h.end():end]
        apply_idx = block.find("[APPLY](")
        if apply_idx != -1:
            block = block[:apply_idx]
        paras = [p.strip() for p in re.split(r"\n\s*\n", block) if p.strip()]
        url = h.group(2)
        job = {
            "job_id": (re.search(r"-(\d+)$", url) or re.search(r"(\d+)", url)).group(1),
            "title": h.group(1).replace("\\", ""),
            "job_url": url,
            "company": "", "job_type": "", "posted_listing": "", "location": "",
            "salary_text": "", "experience": "", "category": "", "role": "",
            "summary": "", "tags": "", "linkedin_note": "", "source_page": page_no,
        }
        # paragraph 0: "Company hidden\n·Full Time·7 Weeks Ago"
        if paras:
            lines = [l.strip() for l in paras[0].splitlines() if l.strip()]
            meta_line = next((l for l in lines if l.startswith("·")), "")
            comp = [l for l in lines if not l.startswith("·")]
            if comp and comp[0].lower() != "company hidden":
                job["company"] = comp[0]
            bits = [b.strip() for b in meta_line.split("·") if b.strip()]
            if bits:
                job["posted_listing"] = bits[-1] if len(bits) > 1 else ""
                job["job_type"] = bits[0]
        # paragraph 1: "United States$93000 - $122K per year5-10 yrs exp [Marketing](..) [Role](..)"
        if len(paras) > 1:
            p = paras[1]
            first_link = p.find("[")
            head = p[:first_link] if first_link != -1 else p
            links = [t for t, _ in LINK_RE.findall(p[first_link:])] if first_link != -1 else []
            em = EXP_RE.search(head)
            if em:
                job["experience"] = em.group(0).replace("yrs exp", "yrs").strip()
                head = head[: em.start()] + head[em.end():]
            sm = SAL_RE.search(head)
            if sm:
                job["salary_text"] = sm.group(0).strip()
                head = head[: sm.start()] + head[sm.end():]
            job["location"] = head.strip()
            # First link is the category; any further links are roles ("Virtual Assistant" is both).
            if links and links[0] in CATEGORIES:
                job["category"], roles = links[0], links[1:]
            else:
                roles = links
            job["role"] = ", ".join(roles)
        for p in paras[2:]:
            if "LinkedIn" in p and ("Not on LinkedIn" in p or "linkedin" in p.lower()) and p.startswith("!"):
                job["linkedin_note"] = strip_md(p)
            elif p.startswith("["):
                job["tags"] = ", ".join(t for t, _ in LINK_RE.findall(p))
            elif not job["summary"] and not p.startswith(("#", "!")):
                job["summary"] = strip_md(p)
        jobs.append(job)
    return jobs


def site_total(md: str):
    m = re.search(r"^#\s+Remote Jobs[^\n]*\((\d[\d,]*)\)", md, re.M)
    return int(m.group(1).replace(",", "")) if m else None


# ---------------------------------------------------------------- detail pages

def parse_detail(md: str, title_meta: str) -> dict:
    d = {"company": "", "company_url": "", "posted_detail": "", "apply_url": "", "description": ""}
    m = re.search(r"\*\*([^*]+)\*\*\]\((https://dailyremote\.com/remote-company/[^)]+)\)", md)
    if m:
        d["company"], d["company_url"] = m.group(1).strip(), m.group(2)
    elif " at " in title_meta:
        d["company"] = title_meta.rsplit(" at ", 1)[1].strip()
    m = re.search(r"^Posted (.+)$", md, re.M)
    if m:
        d["posted_detail"] = m.group(1).strip()
    m = re.search(r"\((https://dailyremote\.com/apply/\d+[^)]*)\)", md)
    if m:
        d["apply_url"] = m.group(1)
    end = md.find("Share this job")
    start = md.find("**Description**")
    if start == -1:
        ai = md.find("AI Summary")
        if ai != -1:
            nxt = md.find("\n\n", md.find("\n\n", ai) + 2)   # skip the summary paragraph
            start = nxt if nxt != -1 else ai
        else:
            # Some pages have no "AI Summary"/"Description" headings: the posting text starts
            # right after the resume-match widget, which ends with a lone "×" line.
            rv = md.find("Get professional review")
            if rv != -1:
                x = md.find("\n×\n", rv)
                start = x + 3 if x != -1 else rv
    else:
        start += len("**Description**")
    if start != -1:
        d["description"] = strip_md(md[start: end if end > start else len(md)])
    return d


# ---------------------------------------------------------------- output

COLUMNS = [
    ("job_id", "Job ID", 11), ("title", "Title", 40), ("company", "Company", 26),
    ("location", "Location", 18), ("job_type", "Job Type", 14), ("experience", "Experience", 11),
    ("salary_text", "Salary (as listed)", 24), ("salary_min", "Salary Min", 12),
    ("salary_max", "Salary Max", 12), ("salary_period", "Salary Period", 10), ("currency", "Currency", 9),
    ("category", "Category", 14), ("role", "Role", 22), ("posted_listing", "Posted", 14),
    ("posted_est", "Posted (est. date)", 13), ("summary", "AI Summary", 60), ("tags", "Tags", 35),
    ("linkedin_note", "LinkedIn Note", 25), ("job_url", "Job URL", 30), ("apply_url", "Apply URL", 30),
    ("company_url", "Company URL", 30), ("description", "Full Description", 80),
    ("source_page", "Source Page", 8), ("status", "Status", 13), ("applied_date", "Applied Date", 13),
    ("notes", "Notes", 30),
]
URL_COLS = {"job_url", "apply_url", "company_url"}


def clean(v):
    if isinstance(v, str):
        v = ILLEGAL_CHARACTERS_RE.sub("", v)
        return v[:32000]
    return v


def write_workbook(rows: list[dict], log: list[tuple], out: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Jobs"
    ws.append([c[1] for c in COLUMNS])
    for r in rows:
        ws.append([clean(r.get(k, "")) for k, _, _ in COLUMNS])
    hdr_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = hdr_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for idx, (key, _, width) in enumerate(COLUMNS, start=1):
        col = ws.cell(row=1, column=idx).column_letter
        ws.column_dimensions[col].width = width
        for row in range(2, ws.max_row + 1):
            c = ws.cell(row=row, column=idx)
            c.alignment = Alignment(vertical="top", wrap_text=key in {"summary", "description", "title"})
            if key in URL_COLS and c.value:
                c.hyperlink = c.value
                c.font = Font(color="0563C1", underline="single")
            if key in {"salary_min", "salary_max"} and c.value is not None:
                c.number_format = "#,##0"
            if key == "posted_est" and c.value:
                c.number_format = "yyyy-mm-dd"
    for row in range(2, ws.max_row + 1):
        ws.row_dimensions[row].height = 60
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions
    status_col = [i for i, c in enumerate(COLUMNS, 1) if c[0] == "status"][0]
    letter = ws.cell(row=1, column=status_col).column_letter
    dv = DataValidation(type="list", formula1='"' + ",".join(STATUS_OPTIONS) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{letter}2:{letter}{max(ws.max_row, 2) + 500}")

    lg = wb.create_sheet("Run Log")
    lg.append(["Field", "Value"])
    for cell in lg[1]:
        cell.font = Font(bold=True)
    for k, v in log:
        lg.append([k, clean(v) if isinstance(v, str) else v])
    lg.column_dimensions["A"].width = 30
    lg.column_dimensions["B"].width = 90
    wb.save(out)


# ---------------------------------------------------------------- main

def main() -> int:
    parser = argparse.ArgumentParser(description="Scrape DailyRemote search results into Excel.")
    parser.add_argument("--search", default="social media", help="Search term on dailyremote.com")
    parser.add_argument("--pages", type=int, default=23, help="Number of listing pages to scrape")
    parser.add_argument("--out", default=None, help="Output .xlsx path (default: project root)")
    parser.add_argument("--workers", type=int, default=3, help="Parallel Firecrawl requests")
    parser.add_argument("--skip-details", action="store_true", help="Don't open each job's page")
    parser.add_argument("--max-details", type=int, default=None, help="Only open the first N job pages")
    parser.add_argument("--limit", type=int, default=None, help="Keep only the first N unique jobs in the sheet")
    args = parser.parse_args()

    api_key = os.environ.get("FIRECRAWL_API_KEY")
    if not api_key:
        print("FIRECRAWL_API_KEY missing from .env", file=sys.stderr)
        return 1
    RAW.mkdir(parents=True, exist_ok=True)
    today = date.today()
    slug = re.sub(r"[^a-z0-9]+", "_", args.search.lower()).strip("_")
    out = Path(args.out) if args.out else ROOT / f"dailyremote_{slug}_jobs_{today.isoformat()}.xlsx"
    started = datetime.now()

    # 1. listing pages
    listing_rows, total_on_site, page_errors = [], None, []
    pages_done = 0
    for n in range(1, args.pages + 1):
        url = f"{BASE}/?search={quote_plus(args.search)}&page={n}"
        try:
            page = fetch(url, RAW / f"{slug}_page_{n:02d}.json", api_key)
        except Exception as e:
            page_errors.append(f"page {n}: {e}")
            print(f"[listing] page {n} FAILED: {e}", file=sys.stderr)
            continue
        jobs = parse_listing(page["markdown"], n)
        if total_on_site is None:
            total_on_site = site_total(page["markdown"])
        pages_done += 1
        print(f"[listing] page {n}: {len(jobs)} jobs")
        if not jobs:
            print(f"[listing] page {n} has no jobs; stopping pagination")
            break
        listing_rows.extend(jobs)

    parsed = len(listing_rows)
    seen, rows = set(), []
    for r in listing_rows:
        if r["job_id"] not in seen:
            seen.add(r["job_id"])
            rows.append(r)
    dupes = parsed - len(rows)
    unique_found = len(rows)
    if args.limit:
        rows = rows[: args.limit]
    if not rows:
        print("No jobs parsed.", file=sys.stderr)
        return 1

    # 2. detail pages
    detail_errors = []
    targets = [] if args.skip_details else rows[: args.max_details] if args.max_details else rows
    if targets:
        def work(r):
            page = fetch(r["job_url"], RAW / f"job_{r['job_id']}.json", api_key)
            return r, parse_detail(page["markdown"], page.get("title", ""))

        with ThreadPoolExecutor(max_workers=args.workers) as ex:
            futs = {ex.submit(work, r): r for r in targets}
            for i, f in enumerate(as_completed(futs), 1):
                r = futs[f]
                try:
                    _, d = f.result()
                    r["company"] = d["company"] or r["company"]
                    r.update({k: d[k] for k in ("company_url", "apply_url", "description")})
                    if d["posted_detail"]:
                        r["posted_detail"] = d["posted_detail"]
                except Exception as e:
                    detail_errors.append(f"{r['job_id']} ({r['job_url']}): {e}")
                if i % 25 == 0 or i == len(targets):
                    print(f"[detail] {i}/{len(targets)} done, {len(detail_errors)} failed")

    # 3. derived fields
    for r in rows:
        r.update(parse_salary(r["salary_text"]))
        # detail page "Posted 2 days ago" is fresher than the listing's age
        r["posted_est"] = estimate_date(r.get("posted_detail") or r["posted_listing"], today)
        r.setdefault("status", "Not applied")
        r["status"] = "Not applied"

    # 4. output
    log = [
        ("Scraped at", started.strftime("%Y-%m-%d %H:%M")),
        ("Search term", args.search),
        ("Source URL", f"{BASE}/?search={quote_plus(args.search)}"),
        ("Listing pages requested", args.pages),
        ("Listing pages scraped", pages_done),
        ("Jobs reported by site", total_on_site),
        ("Rows parsed from listings", parsed),
        ("Duplicates removed", dupes),
        ("Unique jobs found", unique_found),
        ("Row limit (--limit)", args.limit or "none"),
        ("Jobs in sheet", len(rows)),
        ("Gap vs site count (unique found)", (total_on_site - unique_found) if total_on_site else "n/a"),
        ("Job pages opened", len(targets)),
        ("Job pages failed", len(detail_errors)),
        ("Firecrawl credits used this run (cache hits are free)", credits_used),
        ("Listing page errors", "; ".join(page_errors) or "none"),
        ("Job page errors", "; ".join(detail_errors[:50]) or "none"),
        ("Note", "Posted (est. date) is approximate: derived from relative ages like '3 weeks ago'."),
    ]
    write_workbook(rows, log, out)
    csv_path = WORK_DIR / f"{slug}_jobs.csv"
    with csv_path.open("w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow([c[1] for c in COLUMNS])
        for r in rows:
            w.writerow([r.get(k, "") if r.get(k) is not None else "" for k, _, _ in COLUMNS])

    print(f"ok: {len(rows)} jobs -> {out}")
    print(f"csv: {csv_path}")
    print(f"site reports {total_on_site}; duplicates removed {dupes}; job pages failed {len(detail_errors)}; credits used {credits_used}")
    if targets and len(detail_errors) > 0.1 * len(targets):
        print("More than 10% of job pages failed.", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
