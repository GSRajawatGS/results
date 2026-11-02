"""
marksheet_fetcher.py - roll numbers of a RESULT NOTIFICATION Excel -> marksheet data from
https://univindia.com/JwgResult/adselcolclassnew.php  (reuses the browser code of result_scraper.py).

For every course heading of the notification it
  1. reads the site's list of declared exams for that year (cached for a few hours),
  2. picks the BEST MATCHING exam by name (ignores ATKT / NEP / II-chance / month-year wording and
     word order, copes with M.A. = MA, SECOND SEM. = 2 SEM, spelling differences, extra words),
  3. opens that exam's "View Result" page and fetches the marksheet of every roll number,
  4. if a roll is not found, tries the next best matching exams (II chance, ATKT exams ...),
  5. if the exam still looks wrong (no match, or most rolls not found) it ASKS YOU in the terminal which exam
     is the right one (number from the list, exam code, or words to search); the choice is remembered in
     overrides.json (and Supabase) and used from then on - also by the Streamlit app's "Fix marksheets" page,
  6. never repeats work: rolls with a marksheet are skipped, and a marksheet already fetched for the same roll
     and exam in another notification file is reused instead of fetched again,
  7. writes  <notification>_MARKSHEETS.xlsx  (one row per roll, one column per marksheet field).

Needs next to it: result_scraper.py + result_jobs.py (START_URL etc.) and  pip install playwright pandas openpyxl.

Stand-alone use:
  python marksheet_fetcher.py 2026\\SomeNotification.xlsx [--retry] [--headless] [--no-ask]
  python marksheet_fetcher.py list 2026        (re-read the site's list of declared exams)
  python marksheet_fetcher.py match 2026 "M.A. YOGA SECOND SEMESTER (NEP) EXAM.JUNE2026"   (shows the best matches)
"""
import json
import os
import random
import re
import sys
import time
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

MIN_SCORE = 0.60          # a match below this is not used (the roll numbers are reported as NO MATCH)
TRY_WITHIN = 0.15         # next-best exams are also tried when their score is within this of the best
MAX_CANDIDATES = 3        # exams tried per course, best first
LIST_CACHE_HOURS = 12     # how long the site's exam list is trusted

FIXED_COLS = ["Roll No", "Sr No", "Notification Course", "Notification Result", "Result Date", "Matched Exam",
              "Exam Code", "Match Score", "Marksheet Status", "Note", "Fetched At"]


# --------------------------------------------------------------------------
# Exam-name matching
# --------------------------------------------------------------------------
ORD = {"FIRST": 1, "SECOND": 2, "THIRD": 3, "FOURTH": 4, "FIFTH": 5, "SIXTH": 6, "SEVENTH": 7, "EIGHTH": 8,
       "NINTH": 9, "TENTH": 10}
ROMAN = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10}
SUFFIX = {"ST", "ND", "RD", "TH"}
MONTHS = {"JAN": "JAN", "JANUARY": "JAN", "FEB": "FEB", "FEBRUARY": "FEB", "MAR": "MAR", "MARCH": "MAR",
          "APR": "APR", "APRIL": "APR", "MAY": "MAY", "JUN": "JUN", "JUNE": "JUN", "JUL": "JUL", "JULY": "JUL",
          "AUG": "AUG", "AUGUST": "AUG", "SEP": "SEP", "SEPT": "SEP", "SEPTEMBER": "SEP", "OCT": "OCT",
          "OCTOBER": "OCT", "NOV": "NOV", "NOVEMBER": "NOV", "DEC": "DEC", "DECEMBER": "DEC"}
STOP = {"AND", "OF", "THE", "IN", "FOR", "WITH", "EXAM", "EXAMS", "EXAMINATION", "EXAMINATIONS", "A", "AN"}
UNIT_SEM = {"SEM", "SEMESTER", "SMESTER", "SEMS"}
UNIT_YEAR = {"YEAR", "YR"}
# dotted abbreviations: M.A.  M.SC.  A.T.K.T.  B.SC  PH.D.  ->  MA  MSC  ATKT  BSC  PHD
ABBR = re.compile(r"\b(?:[A-Z]{1,2}\.)+(?:[A-Z]{1,3}\b\.?)?")


def parse_exam(text):
    """Break an exam name into: core words, (number, SEM|YEAR), (month, year), ATKT / NEP / chance flags."""
    s = str(text or "").upper()
    s = re.sub(r"\[[^\]]*\]", " ", s).replace("&", " AND ")
    s = ABBR.sub(lambda m: m.group(0).replace(".", "") + " ", s)
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    s = re.sub(r"(?<=[A-Z])(?=\d)|(?<=\d)(?=[A-Z])", " ", s)           # JUNE2026 -> JUNE 2026, 1SEM -> 1 SEM
    raw = s.split()
    toks = []
    for i, t in enumerate(raw):                                          # 1ST / 2ND -> 1 / 2
        if t in SUFFIX and i > 0 and raw[i - 1].isdigit():
            continue
        toks.append(t)

    def num_of(t):
        if t in ORD:
            return ORD[t]
        if t in ROMAN:
            return ROMAN[t]
        if t.isdigit() and 1 <= int(t) <= 12:
            return int(t)
        return None

    used = [False] * len(toks)
    sem = chance = mon = yr = None
    atkt = nep = False
    for i, t in enumerate(toks):
        if t in UNIT_SEM or t in UNIT_YEAR:
            unit = "YEAR" if t in UNIT_YEAR else "SEM"
            n, at = None, None
            if i > 0 and not used[i - 1] and num_of(toks[i - 1]) is not None:
                n, at = num_of(toks[i - 1]), i - 1
            elif i + 1 < len(toks) and num_of(toks[i + 1]) is not None:
                n, at = num_of(toks[i + 1]), i + 1
            used[i] = True
            if n is not None and sem is None:
                sem = (n, unit)
                used[at] = True
        elif t == "CHANCE":
            used[i] = True
            if i > 0 and num_of(toks[i - 1]) is not None:
                chance = num_of(toks[i - 1])
                used[i - 1] = True
            else:
                chance = chance or 2
        elif t == "ATKT":
            atkt, used[i] = True, True
        elif t == "NEP":
            nep, used[i] = True, True
        elif t in MONTHS:
            mon, used[i] = MONTHS[t], True
            if i + 1 < len(toks) and re.fullmatch(r"(19|20)\d\d", toks[i + 1]):
                yr, used[i + 1] = int(toks[i + 1]), True
        elif t in STOP:
            used[i] = True
    for i, t in enumerate(toks):                                         # a year on its own
        if not used[i] and re.fullmatch(r"(19|20)\d\d", t):
            yr, used[i] = (yr or int(t)), True
    core = [t for i, t in enumerate(toks) if not used[i]]
    return {"core": core, "sem": sem, "chance": chance, "session": (mon, yr) if (mon or yr) else None,
            "atkt": atkt, "nep": nep}


def _soft_overlap(a, b):
    left, hit = list(b), 0.0
    for x in a:
        best, bi = 0.0, -1
        for i, y in enumerate(left):
            r = 1.0 if x == y else SequenceMatcher(None, x, y).ratio()
            if r > best:
                best, bi = r, i
        if best >= 0.84:
            hit += best
            del left[bi]
    return hit


def date_key(text):
    m = re.search(r"(\d{1,2})[-./ ](\d{1,2})[-./ ](\d{4})", str(text or ""))
    return (int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None


def score_names(a, b, date_a=None, date_b=None):
    """0..1: how well exam name `b` (site) fits `a` (notification). Both are parse_exam() results."""
    if not a["core"] or not b["core"]:
        return 0.0
    dice = 2 * _soft_overlap(a["core"], b["core"]) / (len(a["core"]) + len(b["core"]))
    ratio = SequenceMatcher(None, " ".join(sorted(a["core"])), " ".join(sorted(b["core"]))).ratio()
    s = 0.6 * dice + 0.4 * ratio
    if a["sem"] and b["sem"]:
        s *= 1.0 if a["sem"] == b["sem"] else 0.3        # a different semester is a different exam
    elif a["sem"] or b["sem"]:
        s *= 0.9
    if a["session"] and b["session"]:
        (m1, y1), (m2, y2) = a["session"], b["session"]
        if (y1 and y2 and y1 != y2) or (m1 and m2 and m1 != m2):
            s *= 0.6
    if a["atkt"] != b["atkt"]:
        s *= 0.92                                         # ATKT wording may be missing: only a small preference
    if (a["chance"] or 1) != (b["chance"] or 1):
        s *= 0.9
    if a["nep"] != b["nep"]:
        s *= 0.9
    if date_a and date_b and date_a == date_b:
        s += 0.05                                         # declared on the same day
    return round(min(s, 1.0), 4)


def rank_exams(exams, name, result_date=None):
    """The site's exams [{'code','name','date'}] sorted best match first, each with a 'score'."""
    a, da = parse_exam(name), date_key(result_date)
    out = []
    for e in exams:
        sc = score_names(a, parse_exam(e.get("name")), da, date_key(e.get("date")))
        out.append({**e, "score": sc})
    out.sort(key=lambda e: (-e["score"], str(e.get("code"))))
    return out


def course_year_session(course, result_date):
    p = parse_exam(course)
    mon, yr = p["session"] or (None, None)
    if not yr:
        d = date_key(result_date)
        yr = d[2] if d else datetime.now().year
    return int(yr), (f"{mon}-{yr}" if mon else "")


# --------------------------------------------------------------------------
# The site (uses the proven code of result_scraper.py)
# --------------------------------------------------------------------------
def _rs():
    try:
        import result_scraper as rs
    except Exception as e:
        raise RuntimeError(f"result_scraper.py / result_jobs.py could not be imported ({type(e).__name__}: {e}). "
                           f"Keep them in the same folder as this script.")
    return rs


def exam_list(p, year, folder, force=False):
    """[{'date','code','name'}] of the declared results of `year` (cached in <folder>/exam_list_<year>.json)."""
    rs = _rs()
    cache = Path(folder) / f"exam_list_{year}.json"
    if cache.exists() and not force and (time.time() - cache.stat().st_mtime) < LIST_CACHE_HOURS * 3600:
        try:
            return json.loads(cache.read_text(encoding="utf-8"))
        except Exception:
            pass
    rs.use_session("", folder, folder, year)
    print(f"   reading the list of declared results for {year} from the site ...")
    rows = rs.read_exam_list(p)
    if not rows:
        raise RuntimeError(f"the site's list for {year} came back empty")
    cache.write_text(json.dumps(rows, indent=1, ensure_ascii=False), encoding="utf-8")
    print(f"   {len(rows)} exams in the list")
    return rows


class _Browser:
    def __init__(self, p, headless):
        self.p, self.headless = p, headless
        self.browser = self.ctx = self.page = None
        self.dialogs = []

    def get(self):
        rs = _rs()
        if self.browser is None:
            self.browser = self.p.chromium.launch(headless=self.headless)
            self.ctx = (self.browser.new_context() if self.headless else self.browser.new_context(no_viewport=True))
            self.dialogs = []
            rs.hook_dialogs(self.ctx, self.dialogs)
            self.page = self.ctx.new_page()
        elif self.page is None or self.page.is_closed():
            self.page = self.ctx.new_page()
        return self.ctx, self.page, self.dialogs

    def close(self):
        try:
            if self.browser:
                self.browser.close()
        except Exception:
            pass
        self.browser = self.ctx = self.page = None


# --------------------------------------------------------------------------
# Output Excel
# --------------------------------------------------------------------------
class Store:
    def __init__(self, path):
        self.path = Path(path)
        self.rows, self.raw = {}, {}
        self.dirty = False
        if self.path.exists():
            try:
                df = pd.read_excel(self.path, sheet_name="Marksheets", dtype=str).fillna("")
                for r in df.to_dict("records"):
                    self.rows[f"{r['Roll No']}|{r['Sr No']}"] = r
                try:
                    rw = pd.read_excel(self.path, sheet_name="Raw", dtype=str).fillna("")
                    self.raw = dict(zip(rw["Roll No"], rw["Marksheet text"]))
                except Exception:
                    pass
            except Exception as e:
                print(f"   (could not read the old {self.path.name}: {e}; starting again)")

    @staticmethod
    def key(roll, sr):
        return f"{roll}|{sr}"

    def done(self, roll, sr, retry_not_found, ov_code=None):
        """True when this roll needs no more work. A NOT FOUND is looked at again when the human has since
        chosen another exam for the course (ov_code), or with --retry."""
        r = self.rows.get(self.key(roll, sr), {})
        st = r.get("Marksheet Status", "")
        if st == "OK":
            return True
        if st == "SKIPPED":
            return not ov_code and not retry_not_found
        if st == "NOT FOUND":
            if ov_code and r.get("Exam Code") != ov_code:
                return False
            return not retry_not_found
        return False

    def put(self, roll, sr, rec, raw=""):
        """Store a row. A row that only differs by its time stamp is left alone, so a run that finds nothing new
        does not change the file (and nothing is uploaded again)."""
        k = self.key(roll, sr)
        old = self.rows.get(k)
        same = lambda a, b: ({x: y for x, y in a.items() if x != "Fetched At"} ==
                             {x: y for x, y in b.items() if x != "Fetched At"})
        if old is None or not same(old, rec):
            self.rows[k] = rec
            self.dirty = True
        if raw and self.raw.get(str(roll)) != raw[:30000]:
            self.raw[str(roll)] = raw[:30000]
            self.dirty = True

    def save(self):
        recs = list(self.rows.values())
        if not recs or (not self.dirty and self.path.exists()):
            return
        extra = []
        for r in recs:
            for k in r:
                if k not in FIXED_COLS and k not in extra:
                    extra.append(k)
        df = pd.DataFrame(recs).reindex(columns=FIXED_COLS + extra).fillna("")
        while True:
            try:
                with pd.ExcelWriter(self.path, engine="openpyxl") as xw:
                    df.to_excel(xw, sheet_name="Marksheets", index=False)
                    if self.raw:
                        pd.DataFrame({"Roll No": list(self.raw), "Marksheet text": list(self.raw.values())}).to_excel(
                            xw, sheet_name="Raw", index=False)
                    ws = xw.book["Marksheets"]
                    ws.freeze_panes = "B2"
                    for col in ws.columns:
                        w = max(len(str(c.value or "")) for c in col[:200])
                        ws.column_dimensions[col[0].column_letter].width = min(max(w + 2, 10), 45)
                self.dirty = False
                return
            except PermissionError:
                if sys.stdin and sys.stdin.isatty():
                    input(f"   '{self.path.name}' is open in Excel. Close it, then press Enter to save... ")
                else:                                   # started from the app: nobody can press Enter - wait, then give up
                    waited = getattr(self, "_waited", 0) + 5
                    self._waited = waited
                    if waited > 120:
                        raise
                    print(f"   '{self.path.name}' is open in Excel - close it (waiting, {waited} s) ...")
                    time.sleep(5)


# --------------------------------------------------------------------------
# Fetching
# --------------------------------------------------------------------------
def _fetch_with_exam(br, rs, cand, rolls, on_found, on_error):
    """Look every roll number up in the exam `cand`. Returns the rolls still without a marksheet."""
    vcls = {"code": cand["code"], "atkt": False, "sem": 0}
    todo = list(rolls)
    total = len(todo)
    for attempt in range(1, 4):
        try:
            ctx, page, dialogs = br.get()
            entry = rs.go_to_entry(ctx, page, vcls)
            br.page = entry
            errors = 0
            for roll in list(todo):
                try:
                    entry = rs.ensure_entry(ctx, entry, vcls)
                    status, note, fields, raw = rs.fetch_one(ctx, entry, roll, dialogs)
                except (rs.BrowserLost, rs.ConfigError):
                    raise
                except Exception as e:
                    if entry.is_closed():
                        raise rs.BrowserLost(str(e))
                    status, note, fields, raw = "ERROR", f"{type(e).__name__}: {str(e)[:150]}", {}, ""
                    try:
                        entry = rs.go_to_entry(ctx, entry, vcls)
                    except Exception:
                        pass
                br.page = entry
                n = total - len(todo) + 1
                tag = f"[{cand['code']}] [{n}/{total}]"
                if status == "OK":
                    on_found(roll, fields, raw, cand)
                    todo.remove(roll)
                    errors = 0
                    print(f"   {tag} {roll}: OK  {'' if os.environ.get('JW_QUIET') else fields.get('Marksheet Name', '')}  {fields.get('Final Total', '')}  "
                          f"{fields.get('Final Result', '')}".rstrip())
                elif status == "NOT FOUND":
                    errors = 0
                    print(f"   {tag} {roll}: not found here")
                else:
                    errors += 1
                    on_error(roll, note, cand)
                    todo.remove(roll)
                    print(f"   {tag} {roll}: ERROR {note}")
                    if errors >= 5:
                        raise rs.ConfigError("5 errors in a row - stopping this exam")
                time.sleep(random.uniform(*rs.DELAY))
            return todo
        except rs.CodeMissing:
            print(f"   {cand['code']} is not on the site's list any more - skipped")
            return todo
        except rs.ConfigError as e:
            print(f"   >>> {str(e)[:200]}")
            return todo
        except Exception as e:                       # BrowserLost / Playwright errors: restart the browser
            print(f"   >>> browser problem ({type(e).__name__}: {str(e)[:100]}); restarting (attempt {attempt}) ...")
            br.close()
    return todo


# --------------------------------------------------------------------------
# The human's choice, and reuse of earlier work
# --------------------------------------------------------------------------
ASK_FRACTION = 0.5      # ask when at least this share of a course's roll numbers has no marksheet
_ASK = {"on": None}     # becomes False when the human types q


def course_key(course, result_date=""):
    """Stable key of a course heading: year | month | course words | semester | NEP  (same key = same exam)."""
    p = parse_exam(course)
    year, _ = course_year_session(course, result_date)
    mon = (p["session"] or (None, None))[0] or ""
    sem = f"{p['sem'][0]}{p['sem'][1][0]}" if p["sem"] else ""
    return f"{year}|{mon}|{' '.join(p['core'])}|{sem}|{'NEP' if p['nep'] else ''}"


def interactive(ask):
    if _ASK["on"] is False:
        return False
    if ask is None:
        return bool(sys.stdin and sys.stdin.isatty())
    return bool(ask)


def _show(lst):
    for i, c in enumerate(lst, 1):
        print(f"      {i:2}. {str(c['code']):<10} {c.get('score', 0):.2f}  {' '.join(str(c['name']).split())[:85]}")


def ask_exam_cli(exams, course, ranked, n, reason):
    """Ask the human which exam is right. -> ('pick', exam) | ('skip', None) | ('later', None) | ('quit', None)"""
    print(f"\n   >>> {reason}")
    print(f"   Course: {course}   ({n} roll number(s) without a marksheet)")
    shown = ranked[:10]
    _show(shown)
    print("   Type: a number from the list | an exam code | words to search the site's list\n"
          "         s = skip this course for good | n = not now | q = stop asking in this run")
    while True:
        try:
            ans = input("   exam> ").strip()
        except EOFError:
            return "quit", None
        low = ans.lower()
        if not ans:
            continue
        if low == "q":
            return "quit", None
        if low == "n":
            return "later", None
        if low == "s":
            return "skip", None
        if ans.isdigit() and 1 <= int(ans) <= len(shown):
            return "pick", shown[int(ans) - 1]
        by_code = [e for e in exams if str(e["code"]).upper() == ans.upper()]
        if len(by_code) == 1:
            return "pick", {**by_code[0], "score": 1.0}
        if len(by_code) > 1:
            shown = [{**e, "score": 1.0} for e in by_code]
            print("   That code is used by several exams - pick one:")
            _show(shown)
            continue
        words = [w for w in re.split(r"\W+", ans.upper()) if w]
        hits = {(str(e["code"]), str(e["name"])) for e in exams if all(w in str(e["name"]).upper() for w in words)}
        everything = rank_exams(exams, ans)
        shown = ([c for c in everything if (str(c["code"]), str(c["name"])) in hits] or everything)[:10]
        print("   Matches:" if hits else "   Nothing contains all those words; closest names:")
        _show(shown)


def _norm_name(text):
    return " ".join(re.sub(r"[^A-Z0-9]+", " ", str(text or "").upper()).split())


def known_index(folder, exclude):
    """(roll, exam code, exam name) -> (row, file) for every OK marksheet in the OTHER notification files."""
    idx = {}
    for f in sorted(Path(folder).glob("*_MARKSHEETS.xlsx")):
        if f.resolve() == Path(exclude).resolve() or f.name.startswith("~$"):
            continue
        try:
            d = pd.read_excel(f, sheet_name="Marksheets", dtype=str).fillna("")
        except Exception:
            continue
        for r in d.to_dict("records"):
            if r.get("Marksheet Status") == "OK" and r.get("Exam Code") and r.get("Roll No"):
                idx[(r["Roll No"], r["Exam Code"], _norm_name(r.get("Matched Exam")))] = (r, f.name)
    return idx


def process_course(p, br, rs, store, ov_store, counts, folder, sr, course, g, retry, ask, known):
    """All roll numbers of one course heading of the notification."""
    result_date = g["Result Date"].iloc[0]
    notif = dict(zip(g["Roll No"], g["Result"]))
    key = course_key(course, result_date)
    ov = ov_store.get(key)
    ov_code = ov.get("code") if ov and not ov.get("skip") and ov.get("code") else None
    pending = [r for r in g["Roll No"] if not store.done(r, sr, retry, ov_code)]
    print(f"\n===== {sr}. {course}  ({len(g)} roll numbers, {len(pending)} to fetch) =====")
    if not pending:
        return

    def base(roll, status, note, cand=None):
        return {"Roll No": roll, "Sr No": sr, "Notification Course": course,
                "Notification Result": notif.get(roll, ""), "Result Date": result_date,
                "Matched Exam": " ".join(str(cand["name"]).split()) if cand else "",
                "Exam Code": cand["code"] if cand else "",
                "Match Score": cand["score"] if cand else "", "Marksheet Status": status,
                "Note": note, "Fetched At": f"{datetime.now():%Y-%m-%d %H:%M:%S}"}

    if ov and ov.get("skip"):
        print("   skipped: you decided earlier that this course is not fetched (remove it in the app's "
              "'Fix marksheets' page or delete its entry in overrides.json to change that)")
        for r in pending:
            store.put(r, sr, base(r, "SKIPPED", "skipped by choice"))
            counts["SKIPPED"] += 1
        store.save()
        return

    year, sess = course_year_session(course, result_date)
    try:
        exams = exam_list(p, year, folder)
        ranked = rank_exams(exams, course, result_date)
    except Exception as e:
        print(f"   >>> cannot read the exam list of {year}: {e}")
        for r in pending:
            store.put(r, sr, base(r, "NO MATCH", f"exam list not available: {str(e)[:120]}"))
            counts["NO MATCH"] += 1
        store.save()
        return

    for c in ranked[:3]:
        print(f"   candidate {c['score']:.2f}  {c['code']}  {' '.join(str(c['name']).split())[:90]}")
    best = ranked[0] if ranked else None
    auto = []
    if best and best["score"] >= MIN_SCORE:
        auto = [c for c in ranked[:MAX_CANDIDATES] if c["score"] >= MIN_SCORE and c["score"] >= best["score"] - TRY_WITHIN]
    manual = None
    if ov_code:
        hit = (next((e for e in ranked if str(e["code"]) == ov_code and _norm_name(e["name"]) == _norm_name(ov.get("name"))), None)
               or next((e for e in ranked if str(e["code"]) == ov_code), None)
               or {"code": ov_code, "name": ov.get("name") or ov_code, "date": ""})
        manual = {**hit, "score": 1.0, "manual": True}
        print(f"   using the exam you chose earlier: {manual['code']}  {' '.join(str(manual['name']).split())[:80]}")
    cands = ([manual] if manual else []) + [c for c in auto if not (manual and c["code"] == manual["code"]
                                                                   and _norm_name(c["name"]) == _norm_name(manual["name"]))]
    if not manual and len(cands) > 1 and cands[1]["score"] >= cands[0]["score"] - 0.02:
        print("   (note: the next candidate scores almost the same - rolls not found in the first are "
              "looked up in the next)")

    rs.use_session(sess, folder, folder, year)           # session text decides between duplicate rows
    saved = {"n": 0}
    tried = []

    def on_found(roll, fields, raw, cand, note=None):
        if note is None:
            note = "exam chosen manually" if cand.get("manual") else ("" if cand is tried[0] else f"found in {cand['code']}")
        rec = base(roll, "OK", note, cand)
        rec.update(fields)
        store.put(roll, sr, rec, raw)
        counts["OK"] += 1
        saved["n"] += 1
        if saved["n"] % 5 == 0:
            store.save()

    def on_error(roll, note, cand):
        store.put(roll, sr, base(roll, "ERROR", note, cand))
        counts["ERROR"] += 1

    def reuse(cand, rolls):
        left = []
        for roll in rolls:
            hit = known.get((roll, str(cand["code"]), _norm_name(cand["name"])))
            if hit:
                rec, fname = hit
                on_found(roll, {k: v for k, v in rec.items() if k not in FIXED_COLS and v != ""}, "", cand,
                         f"reused from {fname}")
                counts["reused"] += 1
                print(f"   [{cand['code']}] {roll}: reused from {fname} (already fetched)")
            else:
                left.append(roll)
        return left

    remaining = list(pending)

    def run(cand_list):
        nonlocal remaining
        for cand in cand_list:
            if not remaining:
                break
            tried.append(cand)
            remaining = reuse(cand, remaining)
            if not remaining:
                break
            if len(tried) > 1:
                print(f"   --- {len(remaining)} roll number(s) not found yet: trying {cand['code']} ---")
            remaining = _fetch_with_exam(br, rs, cand, remaining, on_found, on_error)

    run(cands)

    # nothing found / the exam looks wrong: ask the human
    skipped, rounds = False, 0
    while remaining and rounds < 6 and interactive(ask) and (
            not tried or len(remaining) >= max(1, ASK_FRACTION * len(g))):
        rounds += 1
        reason = ("no exam of the site's list matches this course" if not tried else
                  f"{len(remaining)} of {len(g)} roll numbers were not found in "
                  f"{', '.join(str(c['code']) for c in tried)} - maybe the wrong exam was chosen")
        action, exam = ask_exam_cli(exams, course, ranked, len(remaining), reason)
        if action == "pick":
            if any(str(c["code"]) == str(exam["code"]) and _norm_name(c["name"]) == _norm_name(exam["name"])
                   for c in tried):
                print(f"   {exam['code']} was already searched and none of these roll numbers are in it - "
                      f"choose another exam (or s to skip, n for not now).")
                continue
            ov_store.set_exam(key, course, str(exam["code"]), str(exam["name"]))
            ov_store.save()
            run([{**exam, "score": 1.0, "manual": True}])
        elif action == "skip":
            ov_store.set_skip(key, course)
            ov_store.save()
            for r in remaining:
                store.put(r, sr, base(r, "SKIPPED", "skipped by choice"))
                counts["SKIPPED"] += 1
            remaining, skipped = [], True
        elif action == "quit":
            _ASK["on"] = False
            break
        else:
            break

    if remaining and not skipped:
        if not tried:
            near = "; ".join(f"{c['code']} = {' '.join(str(c['name']).split())[:60]} ({c['score']:.2f})" for c in ranked[:3])
            print("   NO MATCH on the site's list (not declared there yet, or named very differently). "
                  "Choose the exam in the app's 'Fix marksheets' page, or run again in a terminal to be asked.")
            for r in remaining:
                store.put(r, sr, base(r, "NO MATCH", f"nearest: {near}"))
                counts["NO MATCH"] += 1
        else:
            for r in remaining:
                store.put(r, sr, base(r, "NOT FOUND", "looked in: " + ", ".join(str(c["code"]) for c in tried), tried[0]))
                counts["NOT FOUND"] += 1
    store.save()


def fetch_for_results(results_xlsx, retry_not_found=False, headless=None, ask=None):
    """Marksheets for every roll number of a notification Excel (sheet 'Results' made by jiwaji_results_scraper).
    ask: True / False / None (None = ask the human only when running in a terminal)."""
    from playwright.sync_api import sync_playwright
    import jw_store
    rs = _rs()
    results_xlsx = Path(results_xlsx)
    folder = results_xlsx.parent
    df = pd.read_excel(results_xlsx, sheet_name="Results", dtype=str).fillna("")
    if df.empty:
        print("   no roll numbers in", results_xlsx.name)
        return None
    store = Store(folder / f"{results_xlsx.stem}_MARKSHEETS.xlsx")
    ov_store = jw_store.Overrides.for_folder(folder)
    known = known_index(folder, store.path)
    headless = rs.HEADLESS if headless is None else headless
    counts = {"OK": 0, "reused": 0, "NOT FOUND": 0, "ERROR": 0, "NO MATCH": 0, "SKIPPED": 0}

    with sync_playwright() as p:
        br = _Browser(p, headless)
        try:
            for (sr, course), g in df.groupby(["Sr No", "Course / Exam"], sort=False):
                process_course(p, br, rs, store, ov_store, counts, folder, sr, course, g, retry_not_found, ask, known)
        finally:
            br.close()
            store.save()

    print(f"\n   marksheets: {counts}  ->  {store.path}")
    return store.path


# --------------------------------------------------------------------------
# Stand-alone
# --------------------------------------------------------------------------
def _cached_list(year):
    for f in (HERE / str(year) / f"exam_list_{year}.json", HERE / f"exam_list_{year}.json"):
        if f.exists():
            return json.loads(f.read_text(encoding="utf-8"))
    return None


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = {a for a in sys.argv[1:] if a.startswith("--")}
    if args and args[0].lower() == "match":
        if len(args) < 3:
            sys.exit('Usage: python marksheet_fetcher.py match 2026 "exam name as in the notification"')
        year, name = args[1], args[2]
        exams = _cached_list(year)
        if exams is None:
            from playwright.sync_api import sync_playwright
            (HERE / str(year)).mkdir(exist_ok=True)
            with sync_playwright() as p:
                exams = exam_list(p, int(year), HERE / str(year))
        for c in rank_exams(exams, name)[:8]:
            print(f"{c['score']:.2f}  {c['code']}  {' '.join(str(c['name']).split())}")
        return
    if args and args[0].lower() == "list":
        from playwright.sync_api import sync_playwright
        year = int(args[1]) if len(args) > 1 else datetime.now().year
        (HERE / str(year)).mkdir(exist_ok=True)
        with sync_playwright() as p:
            rows = exam_list(p, year, HERE / str(year), force=True)
        print(f"{len(rows)} exams saved in {HERE / str(year)}")
        return
    if not args:
        sys.exit(__doc__)
    for a in args:
        fetch_for_results(a, retry_not_found="--retry" in flags, headless=True if "--headless" in flags else None,
                          ask=False if "--no-ask" in flags else None)


if __name__ == "__main__":
    main()
