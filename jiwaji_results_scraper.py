"""
Jiwaji University - RESULT NOTIFICATION downloader + roll-number extractor.

What it does
  1. Opens https://jiwaji.edu/notice-board in Chrome
  2. Selects the year (YEAR below) in "Select Year"
  3. Finds every link containing "RESULT NOTIFICATION" (newest first)
  4. Skips links already downloaded, clicks the next new one, saves the PDF to
     ./<YEAR>/<link name>.pdf
  5. Extracts roll numbers + all the information on the same row into Excel
     (./<YEAR>/<link name>.xlsx, one row per roll number, info in next columns)
  6. For every roll number, fetches the MARKSHEET from univindia.com (see marksheet_fetcher.py):
     the best-matching exam is chosen by name, result -> ./<YEAR>/<link name>_MARKSHEETS.xlsx
  7. Uploads the result to Supabase (when SUPABASE_URL / SUPABASE_KEY are in .env) - only what changed
  8. Repeats until no new RESULT NOTIFICATION links remain

At the start it prints what is already done (PDF / Excel / marksheets / uploaded) and finishes unfinished
work first. Nothing is downloaded, extracted, fetched or uploaded twice.

Install:
  pip install selenium pdfplumber pandas openpyxl
  (Selenium 4.6+ downloads the matching chromedriver automatically.)

Run:
  python jiwaji_results_scraper.py            # uses YEAR below
  python jiwaji_results_scraper.py 2025       # or pass a year
  python jiwaji_results_scraper.py 2026 --no-marksheets     # download + roll numbers only
  python jiwaji_results_scraper.py 2026 --marksheets-only   # marksheets for the Excel files already in ./2026
  python jiwaji_results_scraper.py 2026 --marksheets-only --retry   # also look again at NOT FOUND rolls
  python jiwaji_results_scraper.py 2026 --plan              # only show what is done / still to do
  python jiwaji_results_scraper.py 2026 --no-ask            # never ask which exam is right (unattended)
  python jiwaji_results_scraper.py 2026 --no-sync           # do not upload to Supabase

Marksheets need result_scraper.py + result_jobs.py + marksheet_fetcher.py in the same folder,
and  pip install playwright  (then: playwright install chromium).
"""

import argparse
import re
import sys
import time
from pathlib import Path

import pandas as pd
import pdfplumber
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait

URL = "https://jiwaji.edu/notice-board"
_ap = argparse.ArgumentParser(add_help=True)
_ap.add_argument("year", nargs="?", default="2026")
_ap.add_argument("--no-marksheets", action="store_true", help="skip the marksheet step")
_ap.add_argument("--marksheets-only", action="store_true", help="only fetch marksheets for existing Excel files")
_ap.add_argument("--retry", action="store_true", help="look again at roll numbers that were NOT FOUND")
_ap.add_argument("--plan", action="store_true", help="only show what is done and what is still to do")
_ap.add_argument("--no-ask", action="store_true", help="never ask which exam is the right one")
_ap.add_argument("--no-sync", action="store_true", help="do not upload to Supabase")
_ap.add_argument("--force-sync", action="store_true", help="upload again even if unchanged")
_ap.add_argument("--central", action="store_true",
                 help="cloud mode: notifications already in Supabase count as done; only roll numbers are uploaded, "
                      "marksheets are left to the workers (jw_cloud.py)")
_ap.add_argument("--headless", action="store_true", help="Chrome without a window (servers, GitHub Actions)")
ARGS, _ = _ap.parse_known_args()
if ARGS.central:
    ARGS.no_marksheets = True
YEAR = ARGS.year
KEYWORD = "RESULT NOTIFICATION"
BASE_DIR = Path(__file__).resolve().parent
YEAR_DIR = BASE_DIR / YEAR
YEAR_DIR.mkdir(parents=True, exist_ok=True)

# A roll number: 6-14 chars, digits, optionally with letters (tweak if needed)
ROLL_RE = re.compile(r"^(?=.*\d)[A-Za-z0-9]{6,14}$")


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def safe_name(text: str) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    return re.sub(r'[\\/:*?"<>|]', "-", text)[:150] or "result"


def build_driver() -> webdriver.Chrome:
    opts = webdriver.ChromeOptions()
    if ARGS.headless:
        for a in ("--headless=new", "--no-sandbox", "--disable-dev-shm-usage", "--window-size=1400,1000"):
            opts.add_argument(a)
    else:
        opts.add_argument("--start-maximized")
    opts.add_experimental_option(
        "prefs",
        {
            "download.default_directory": str(YEAR_DIR),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,  # download, don't preview
            "safebrowsing.enabled": True,
        },
    )
    return webdriver.Chrome(options=opts)


def select_year(driver) -> None:
    """Click the year button (e.g. 2026) next to 'Select Year:'.
    Matches visible text OR value/title/data attributes, in the main page or any iframe."""
    xp = (
        f"//*[self::a or self::button or self::input or self::span or self::div or self::li or self::label]"
        f"[normalize-space(text())='{YEAR}' or @value='{YEAR}' or @title='{YEAR}' "
        f"or @data-year='{YEAR}' or normalize-space(.)='{YEAR}']"
    )

    def try_click(d) -> bool:
        # prefer the innermost visible match
        for el in reversed(d.find_elements(By.XPATH, xp)):
            try:
                if el.is_displayed():
                    d.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                    d.execute_script("arguments[0].click();", el)
                    return True
            except Exception:
                continue
        return False

    end = time.time() + 30
    while time.time() < end:
        driver.switch_to.default_content()
        if try_click(driver):
            time.sleep(3)
            return
        for fr in driver.find_elements(By.TAG_NAME, "iframe"):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(fr)
                if try_click(driver):
                    driver.switch_to.default_content()
                    time.sleep(3)
                    return
            except Exception:
                continue
        driver.switch_to.default_content()
        time.sleep(1)
    raise RuntimeError(f"Could not find a '{YEAR}' button on the page")


def result_links(driver):
    """Return [(link_text, element)] for every RESULT NOTIFICATION link, in page order (newest first)."""
    xp = ("//a[contains(translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyz',"
          "'ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'RESULT NOTIFICATION')]")
    out = []
    for el in driver.find_elements(By.XPATH, xp):
        txt = re.sub(r"\s+", " ", el.text or el.get_attribute("title") or "").strip()
        if txt:
            out.append((txt, el))
    return out


def wait_for_download(before: set, timeout=90) -> Path | None:
    end = time.time() + timeout
    while time.time() < end:
        pdfs = {p for p in YEAR_DIR.glob("*.pdf")} - before
        partial = list(YEAR_DIR.glob("*.crdownload"))
        if pdfs and not partial:
            return max(pdfs, key=lambda p: p.stat().st_mtime)
        time.sleep(1)
    return None


def download_link(driver, name: str, el) -> Path | None:
    target = YEAR_DIR / f"{safe_name(name)}.pdf"
    before = set(YEAR_DIR.glob("*.pdf"))
    main = driver.current_window_handle
    handles_before = set(driver.window_handles)

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    driver.execute_script("arguments[0].click();", el)

    time.sleep(2)
    # if the link opened a new tab, close it (the PDF is downloaded by Chrome)
    for h in set(driver.window_handles) - handles_before:
        driver.switch_to.window(h)
        driver.close()
    driver.switch_to.window(main)

    got = wait_for_download(before)
    if not got:
        return None
    if got != target:
        if target.exists():
            target.unlink()
        got.rename(target)
    return target


# --------------------------------------------------------------------------
# PDF -> Excel
# --------------------------------------------------------------------------
# Layout of these PDFs:
#   "<n>. <COURSE / EXAM NAME>"      <- heading (English, ASCII)
#   "PASS" / "FAIL" / "A.T.K.T."     <- result status line
#   "251001216 251001217 ..."        <- roll numbers (one or more lines)
HEADING_RE = re.compile(r"^(\d{1,3})\.\s+([A-Za-z][\x20-\x7e]+)$")
STATUS_RE = re.compile(r"^[A-Z][A-Z .\-/]{1,28}$")
ROLL_LINE_RE = re.compile(r"^\d{6,}(\s+\d{6,})*$")
DATE_RE = re.compile(r"(\d{2})[-.](\d{2})[-.](\d{4})")


def parse_result_pdf(pdf_path: Path) -> list[dict]:
    m = DATE_RE.search(pdf_path.stem)
    result_date = f"{m.group(1)}-{m.group(2)}-{m.group(3)}" if m else ""

    records, sr, course, status = [], "", None, None
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for raw in (page.extract_text() or "").splitlines():
                line = re.sub(r"\s+", " ", raw).strip()
                if not line:
                    continue
                if line.startswith("~~~~"):          # end of result section
                    course = status = None
                    continue
                h = HEADING_RE.match(line)
                if h:
                    sr, course, status = h.group(1), h.group(2).strip(), None
                    continue
                if course and ROLL_LINE_RE.match(line):
                    for roll in line.split():
                        records.append({
                            "Roll No": roll,
                            "Result": status or "",
                            "Sr No": sr,
                            "Course / Exam": course,
                            "Result Date": result_date,
                            "Source PDF": pdf_path.name,
                        })
                    continue
                if course and STATUS_RE.match(line):
                    status = line
    return records


def pdf_to_excel(pdf_path: Path) -> Path:
    records = parse_result_pdf(pdf_path)
    cols = ["Roll No", "Result", "Sr No", "Course / Exam", "Result Date", "Source PDF"]
    df = pd.DataFrame(records, columns=cols).drop_duplicates()
    summary = (
        df.groupby(["Course / Exam", "Result"]).size().reset_index(name="Count")
        if not df.empty else pd.DataFrame(columns=["Course / Exam", "Result", "Count"])
    )

    out = pdf_path.with_suffix(".xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="Results", index=False)
        summary.to_excel(xw, sheet_name="Summary", index=False)
        for ws in xw.book.worksheets:                      # readable column widths
            for col in ws.columns:
                width = max(len(str(c.value or "")) for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(width + 2, 70)
            ws.freeze_panes = "A2"
    print(f"    -> {len(df)} roll numbers written to {out.name}")
    return out


# --------------------------------------------------------------------------
# Marksheets (univindia.com), Supabase, and "what is already done"
# --------------------------------------------------------------------------
LEDGER_FILE = BASE_DIR / "ledger.json"


def get_marksheets(results_xlsx: Path):
    if ARGS.no_marksheets:
        return
    try:
        import marksheet_fetcher
    except Exception as e:
        print(f"    !! marksheet step skipped: marksheet_fetcher.py could not be imported "
              f"({type(e).__name__}: {e})")
        return
    try:
        marksheet_fetcher.fetch_for_results(results_xlsx, retry_not_found=ARGS.retry,
                                            ask=False if ARGS.no_ask else None)
    except Exception as e:
        print(f"    !! marksheet step failed for {results_xlsx.name}: {type(e).__name__}: {e}")


def sync_supabase(results_xlsx: Path):
    """Upload one notification to Supabase (skipped when unchanged since the last upload)."""
    if ARGS.no_sync:
        return
    try:
        import jw_store
        if not jw_store.configured():
            return
        led = jw_store.Ledger(LEDGER_FILE)
        r = jw_store.sync_notification(jw_store.get_client(), results_xlsx, led, force=ARGS.force_sync)
        led.save()
        if r.get("skipped"):
            print("    Supabase: already up to date")
        else:
            print(f"    Supabase: {r.get('results', 0)} roll numbers, {r.get('marksheets', 0)} marksheet rows uploaded")
        try:                                              # also the exam choices and the exam list
            ov = jw_store.Overrides(BASE_DIR / "overrides.json")
            jw_store.push_overrides(jw_store.get_client(), ov)
        except Exception:
            pass
    except Exception as e:
        print(f"    !! Supabase upload skipped ({type(e).__name__}: {str(e)[:150]})")


def show_plan():
    """What is already done in this year's folder."""
    try:
        import jw_store
        led = jw_store.Ledger(LEDGER_FILE)
        table = jw_store.plan(YEAR_DIR, led, need_sync=jw_store.configured() and not ARGS.no_sync)
    except Exception as e:
        print(f"(could not build the overview: {type(e).__name__}: {e})")
        return None
    if table.empty:
        print(f"Nothing downloaded yet in {YEAR_DIR}")
        return table
    print(f"\nWhat is already done in {YEAR_DIR}:")
    t = table.copy()
    t["notification"] = t["notification"].str.slice(0, 48)
    print(t.to_string(index=False))
    print()
    return table


def process_pdf(pdf: Path):
    """PDF -> roll-number Excel -> marksheets -> Supabase. Every step is skipped when it is already done."""
    xlsx = pdf.with_suffix(".xlsx")
    if xlsx.exists() and xlsx.stat().st_mtime >= pdf.stat().st_mtime:
        print(f"    roll numbers already extracted ({xlsx.name})")
    else:
        try:
            xlsx = pdf_to_excel(pdf)
        except Exception as e:
            print(f"    !! extraction failed for {pdf.name}: {e}")
            return
    get_marksheets(xlsx)            # skips rolls that already have a marksheet
    sync_supabase(xlsx)             # skips files that are already uploaded


def unfinished_pdfs():
    """Downloaded PDFs that still need extraction, marksheets or upload."""
    try:
        import jw_store
        led = jw_store.Ledger(LEDGER_FILE)
        t = jw_store.plan(YEAR_DIR, led, need_sync=jw_store.configured() and not ARGS.no_sync)
    except Exception:
        return sorted(YEAR_DIR.glob("*.pdf"))
    if t.empty:
        return []
    if ARGS.retry:
        return [YEAR_DIR / f"{n}.pdf" for n in t["notification"]]
    return [YEAR_DIR / f"{n}.pdf" for n, todo in zip(t["notification"], t["to do"]) if todo != "-"]


def central_done():
    try:
        import jw_store
        rows = jw_store.fetch_all(jw_store.get_client(), "notifications", "id,year")
        return {r["id"] for r in rows if str(r.get("year") or "") == str(YEAR)} if ARGS.year else {r["id"] for r in rows}
    except Exception as e:
        print(f"   (could not ask Supabase which notifications exist: {type(e).__name__}: {str(e)[:100]})")
        return set()


# --------------------------------------------------------------------------
# Main loop
# --------------------------------------------------------------------------
def main():
    show_plan()
    if ARGS.plan:
        return

    # 1) finish what was started earlier (no download needed)
    todo = unfinished_pdfs()
    if todo:
        print(f"{len(todo)} downloaded notification(s) have unfinished work - finishing them first")
    for pdf in todo:
        print(f"\nFinishing: {pdf.stem}")
        process_pdf(pdf)
    if ARGS.marksheets_only:
        return

    # 2) new RESULT NOTIFICATION links
    driver = build_driver()
    try:
        driver.get(URL)
        WebDriverWait(driver, 30).until(lambda d: d.find_elements(By.TAG_NAME, "a"))
        select_year(driver)

        while True:
            links = result_links(driver)
            done = {p.stem for p in YEAR_DIR.glob("*.pdf")} | {p.stem for p in YEAR_DIR.glob("*.failed")}
            if ARGS.central:                                   # already uploaded by this or another computer
                done |= central_done()
            todo_links = [(t, el) for t, el in links if safe_name(t) not in done]

            if not todo_links:
                print("No new RESULT NOTIFICATION links. Finished.")
                break

            text, el = todo_links[0]  # first = most recent
            print(f"Downloading: {text}")
            pdf = download_link(driver, text, el)
            if pdf is None:
                print("    !! download failed/timed out, skipping this link")
                (YEAR_DIR / f"{safe_name(text)}.failed").touch()   # so the loop does not spin forever
                continue

            process_pdf(pdf)
            time.sleep(1)
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
