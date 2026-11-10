"""
jw_store.py - the data layer shared by the scraper and the Streamlit app.

  * reads the Excel files made by jiwaji_results_scraper.py / marksheet_fetcher.py   (local files)
  * or reads the same data from Supabase tables                                       (cloud)
  * uploads (syncs) local data to Supabase, remembering what is already uploaded
  * remembers the exam the human chose for a course (overrides.json  +  table exam_overrides)
  * searches students: name or part of a name (any word order, small typos), roll number,
    enrollment number, father's / mother's name, course

Folder layout (root = the folder of the scripts):
    <root>/2026/<notification>.pdf
    <root>/2026/<notification>.xlsx                 sheet "Results"      (roll numbers of the notification)
    <root>/2026/<notification>_MARKSHEETS.xlsx      sheets "Marksheets", "Raw"
    <root>/2026/exam_list_2026.json                 the site's list of declared exams
    <root>/overrides.json   <root>/ledger.json

Supabase: run supabase_schema.sql once, then put SUPABASE_URL and SUPABASE_KEY (service_role key) in a .env file.

Command line:
    python jw_store.py sync [--force]     upload everything that changed to Supabase
    python jw_store.py counts             rows in every Supabase table
    python jw_store.py status             what is done locally (same table the scraper prints)
"""
import json
import math
import os
import re
import sys
import time
from datetime import datetime
from difflib import get_close_matches
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent


def load_env():
    try:
        from dotenv import load_dotenv
        load_dotenv(HERE / ".env")
    except Exception:
        pass


load_env()

FIXED = ["Roll No", "Sr No", "Notification Course", "Notification Result", "Result Date", "Matched Exam",
         "Exam Code", "Match Score", "Marksheet Status", "Note", "Fetched At"]
NAME_EXCL = ("father", "mother", "college", "exam", "guardian", "centre", "center")
KEYS = ["notification", "sr_no", "roll_no"]
MS_COLS = ["notification", "year", "sr_no", "roll_no", "course", "result", "result_date", "status", "note",
           "matched_exam", "exam_code", "match_score", "fetched_at", "name", "father_name", "mother_name",
           "enrollment_no", "final_result", "final_total", "obtained", "maximum", "percentage", "fields"]
RES_COLS = ["notification", "year", "sr_no", "roll_no", "course", "result", "result_date"]
DONE_STATUSES = ("OK",)


# --------------------------------------------------------------------------
# small helpers
# --------------------------------------------------------------------------
def s(v):
    """Any cell -> clean string ('' for empty / NaN)."""
    if v is None:
        return ""
    if isinstance(v, float) and math.isnan(v):
        return ""
    return str(v).strip()


def to_num(v):
    t = s(v)
    try:
        return float(t) if t else None
    except ValueError:
        return None


def split_marks(text):
    m = re.search(r"(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)", s(text))
    if m:
        return float(m.group(1)), float(m.group(2))
    n = re.search(r"\d+(?:\.\d+)?", s(text))
    return (float(n.group(0)), None) if n else (None, None)


def norm_text(text):
    return " ".join(re.sub(r"[^A-Z0-9]+", " ", s(text).upper()).split())


def clean_json(v):
    """Make a value safe for JSON / Supabase (NaN -> None, numpy -> python)."""
    if v is None:
        return None
    if isinstance(v, dict):
        return {str(k): clean_json(x) for k, x in v.items()}
    if isinstance(v, (list, tuple)):
        return [clean_json(x) for x in v]
    if hasattr(v, "item") and not isinstance(v, (str, bytes)):
        try:
            v = v.item()
        except Exception:
            pass
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return None
    return v


# --------------------------------------------------------------------------
# local files
# --------------------------------------------------------------------------
def year_folders(root):
    root = Path(root)
    if not root.is_dir():
        return []
    return [d for d in sorted(root.iterdir()) if d.is_dir() and re.fullmatch(r"\d{4}", d.name)]


def results_files(folder):
    """Notification Excel files of a year folder (sheet 'Results')."""
    out = []
    for p in sorted(Path(folder).glob("*.xlsx")):
        if p.name.startswith("~$") or p.stem.endswith("_MARKSHEETS"):
            continue
        try:
            if "Results" in pd.ExcelFile(p).sheet_names:
                out.append(p)
        except Exception:
            continue
    return out


def marksheet_path(results_xlsx):
    p = Path(results_xlsx)
    return p.with_name(f"{p.stem}_MARKSHEETS.xlsx")


def read_results(path, year=""):
    df = pd.read_excel(path, sheet_name="Results", dtype=str).fillna("")
    out = pd.DataFrame({
        "notification": Path(path).stem, "year": str(year),
        "sr_no": df.get("Sr No", ""), "roll_no": df.get("Roll No", ""),
        "course": df.get("Course / Exam", ""), "result": df.get("Result", ""),
        "result_date": df.get("Result Date", ""),
    })
    return out[RES_COLS]


def _pick(cols, test):
    return next((c for c in cols if test(c.lower())), None)


def read_marksheets(path, notification, year=""):
    """Marksheets sheet -> standard columns (name, father_name, ..., fields = every other column)."""
    df = pd.read_excel(path, sheet_name="Marksheets", dtype=str).fillna("")
    cols = list(df.columns)
    c_name = _pick(cols, lambda l: l.startswith("marksheet") and "name" in l and not any(x in l for x in NAME_EXCL))
    c_father = _pick(cols, lambda l: l.startswith("marksheet") and "father" in l)
    c_mother = _pick(cols, lambda l: l.startswith("marksheet") and "mother" in l)
    c_enrol = _pick(cols, lambda l: l.startswith("marksheet") and "enrol" in l)
    c_fres = _pick(cols, lambda l: l.startswith("final") and "result" in l)
    c_ftot = _pick(cols, lambda l: l.startswith("final") and "total" in l)
    rows = []
    for r in df.to_dict("records"):
        total = s(r.get(c_ftot)) if c_ftot else ""
        ob, mx = split_marks(total)
        rows.append({
            "notification": notification, "year": str(year), "sr_no": s(r.get("Sr No")), "roll_no": s(r.get("Roll No")),
            "course": s(r.get("Notification Course")), "result": s(r.get("Notification Result")),
            "result_date": s(r.get("Result Date")), "status": s(r.get("Marksheet Status")), "note": s(r.get("Note")),
            "matched_exam": s(r.get("Matched Exam")), "exam_code": s(r.get("Exam Code")),
            "match_score": to_num(r.get("Match Score")), "fetched_at": s(r.get("Fetched At")),
            "name": s(r.get(c_name)) if c_name else "", "father_name": s(r.get(c_father)) if c_father else "",
            "mother_name": s(r.get(c_mother)) if c_mother else "",
            "enrollment_no": s(r.get(c_enrol)) if c_enrol else "",
            "final_result": s(r.get(c_fres)) if c_fres else "", "final_total": total,
            "obtained": ob, "maximum": mx,
            "percentage": round(100 * ob / mx, 2) if ob is not None and mx else None,
            "fields": {k: v for k, v in r.items() if k not in FIXED and s(v)},
        })
    return pd.DataFrame(rows, columns=MS_COLS)


def read_raw(path):
    try:
        raw = pd.read_excel(path, sheet_name="Raw", dtype=str).fillna("")
        return dict(zip(raw["Roll No"], raw["Marksheet text"]))
    except Exception:
        return {}


def combine(results, ms):
    """One row per roll number of every notification, with the marksheet columns added (status NOT FETCHED if none)."""
    if ms is None or ms.empty:
        out = results.copy()
        for c in MS_COLS:
            if c not in out:
                out[c] = None
    else:
        msx = ms.drop(columns=["year", "course", "result", "result_date"], errors="ignore")
        out = results.merge(msx, on=KEYS, how="left")
    out["status"] = out["status"].fillna("").replace("", "NOT FETCHED")
    for c in ("name", "father_name", "mother_name", "enrollment_no", "final_result", "final_total", "exam_code",
              "matched_exam", "note"):
        out[c] = out[c].fillna("")
    return out


def load_local(root):
    res, mss = [], []
    for fy in year_folders(root):
        for rf in results_files(fy):
            res.append(read_results(rf, fy.name))
            mp = marksheet_path(rf)
            if mp.exists():
                try:
                    mss.append(read_marksheets(mp, rf.stem, fy.name))
                except Exception as e:
                    print(f"   (could not read {mp.name}: {e})")
    results = pd.concat(res, ignore_index=True) if res else pd.DataFrame(columns=RES_COLS)
    ms = pd.concat(mss, ignore_index=True) if mss else pd.DataFrame(columns=MS_COLS)
    return {"results": results, "marksheets": ms, "combined": combine(results, ms)}


def local_exam_lists(root):
    rows = []
    for fy in year_folders(root):
        for f in fy.glob("exam_list_*.json"):
            try:
                rows += json.loads(f.read_text(encoding="utf-8"))
            except Exception:
                pass
    return _dedupe_exams(rows)


def _dedupe_exams(rows):
    seen, out = set(), []
    for r in rows:
        k = (s(r.get("code")), s(r.get("name")))
        if k not in seen and k[0]:
            seen.add(k)
            out.append({"code": k[0], "name": k[1], "date": s(r.get("date"))})
    return out


# --------------------------------------------------------------------------
# What is done (avoid doing it twice)
# --------------------------------------------------------------------------
class Ledger:
    """ledger.json: remembers which Excel files are already uploaded to Supabase (by size + time)."""

    def __init__(self, path):
        self.path = Path(path)
        self.data = {"synced": {}}
        if self.path.exists():
            try:
                self.data.update(json.loads(self.path.read_text(encoding="utf-8")))
            except Exception:
                pass

    @staticmethod
    def sig(*files):
        out = []
        for f in files:
            f = Path(f)
            out.append([f.stat().st_size, int(f.stat().st_mtime)] if f.exists() else None)
        return out

    def synced(self, key, *files):
        return self.data["synced"].get(key, {}).get("sig") == self.sig(*files)

    def mark(self, key, *files):
        self.data["synced"][key] = {"sig": self.sig(*files), "at": f"{datetime.now():%Y-%m-%d %H:%M:%S}"}

    def save(self):
        self.path.write_text(json.dumps(self.data, indent=1), encoding="utf-8")


def plan(folder, ledger=None, need_sync=False):
    """One row per downloaded PDF of a year folder: what exists and what is still to do."""
    folder = Path(folder)
    rows = []
    for pdf in sorted(folder.glob("*.pdf")):
        rx = pdf.with_suffix(".xlsx")
        mx = marksheet_path(rx)
        row = {"notification": pdf.stem, "excel": rx.exists(), "rolls": 0, "OK": 0, "NOT FOUND": 0, "NO MATCH": 0,
               "ERROR": 0, "SKIPPED": 0, "not fetched": 0, "synced": "", "to do": ""}
        if rx.exists():
            try:
                res = read_results(rx)
                row["rolls"] = len(res)
                if mx.exists():
                    ms = pd.read_excel(mx, sheet_name="Marksheets", dtype=str).fillna("")
                    st = ms["Marksheet Status"].value_counts().to_dict()
                    for k in ("OK", "NOT FOUND", "NO MATCH", "ERROR", "SKIPPED"):
                        row[k] = int(st.get(k, 0))
                    known = set(zip(ms["Roll No"], ms["Sr No"]))
                else:
                    known = set()
                row["not fetched"] = int(sum(1 for r, sr in zip(res["roll_no"], res["sr_no"]) if (r, sr) not in known))
            except Exception as e:
                row["to do"] = f"unreadable: {type(e).__name__}"
        todo = []
        if not row["excel"]:
            todo.append("extract roll numbers")
        if row["not fetched"] or row["NO MATCH"] or row["ERROR"]:
            todo.append("marksheets")
        if ledger is not None and need_sync and row["excel"]:
            ok = ledger.synced(pdf.stem, rx, mx)
            row["synced"] = "yes" if ok else "no"
            if not ok:
                todo.append("sync")
        row["to do"] = row["to do"] or ", ".join(todo) or "-"
        rows.append(row)
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------
# The exam the human chose for a course
# --------------------------------------------------------------------------
class Overrides:
    """overrides.json: key (year|month|course words|semester|NEP) -> {'code','name','course','skip','decided_at','by'}."""

    def __init__(self, path):
        self.path = Path(path)
        self.data = {}
        if self.path.exists():
            try:
                self.data = json.loads(self.path.read_text(encoding="utf-8"))
            except Exception:
                self.data = {}

    @classmethod
    def for_folder(cls, year_folder):
        return cls(Path(year_folder).parent / "overrides.json")

    def get(self, key):
        return self.data.get(key)

    def set_exam(self, key, course, code, name, by="manual"):
        self.data[key] = {"code": code, "name": name, "course": course, "skip": False,
                          "decided_at": f"{datetime.now():%Y-%m-%d %H:%M:%S}", "by": by}

    def set_skip(self, key, course, by="manual"):
        self.data[key] = {"code": "", "name": "", "course": course, "skip": True,
                          "decided_at": f"{datetime.now():%Y-%m-%d %H:%M:%S}", "by": by}

    def remove(self, key):
        self.data.pop(key, None)

    def save(self):
        self.path.write_text(json.dumps(self.data, indent=1, ensure_ascii=False), encoding="utf-8")


# --------------------------------------------------------------------------
# Supabase
# --------------------------------------------------------------------------
def configured():
    load_env()
    return bool(os.environ.get("SUPABASE_URL") and os.environ.get("SUPABASE_KEY"))


def get_client():
    load_env()
    url, key = os.environ.get("SUPABASE_URL"), os.environ.get("SUPABASE_KEY")
    if not (url and key):
        raise RuntimeError("SUPABASE_URL / SUPABASE_KEY are not set (put them in the .env file).")
    from supabase import create_client
    return create_client(url, key)


def upsert(client, table, rows, on_conflict, batch=400):
    rows = [clean_json(r) for r in rows]
    cols = [c.strip() for c in on_conflict.split(",")]
    last = {}                      # the same key twice in one upload makes Postgres refuse it: keep the last row
    for r in rows:
        last[tuple(r.get(c) for c in cols)] = r
    rows = list(last.values())
    for i in range(0, len(rows), batch):
        client.table(table).upsert(rows[i:i + batch], on_conflict=on_conflict).execute()
    return len(rows)


def fetch_all(client, table, columns="*", order=(), page=1000, where=None):
    """All rows of a table (or of the rows where column = value), page by page. Prefer `where` on big tables."""
    out, start = [], 0
    while True:
        q = client.table(table).select(columns)
        for k, v in (where or {}).items():
            q = q.eq(k, v)
        for o in order:
            q = q.order(o)
        rows = q.range(start, start + page - 1).execute().data or []
        out += rows
        if len(rows) < page:
            return out
        start += page


STORE_RAW_TEXT = os.environ.get("JW_STORE_RAW") == "1"   # the full marksheet text is NOT uploaded unless this is set (saves space)
MS_ONLY_IN_RESULTS = ("year", "course", "result", "result_date")


def _ms_rows(ms, raw):
    rows = []
    for r in ms.to_dict("records"):
        r = {k: (None if (isinstance(v, float) and math.isnan(v)) else v) for k, v in r.items()}
        for k in MS_ONLY_IN_RESULTS:               # these live in notification_results, not in the marksheets table
            r.pop(k, None)
        if STORE_RAW_TEXT:
            r["raw_text"] = raw.get(r["roll_no"])
        rows.append(r)
    return rows


def remote_ok_keys(client, notification, page=1000):
    """(sr_no, roll_no) of the marksheets of this notification that are already finished (status OK) in Supabase."""
    keys, start = set(), 0
    while True:
        r = (client.table("marksheets").select("sr_no,roll_no").eq("notification", notification).eq("status", "OK")
             .range(start, start + page - 1).execute().data or [])
        keys |= {(x["sr_no"], x["roll_no"]) for x in r}
        if len(r) < page:
            return keys
        start += page


def drop_weaker(rows, ok_keys):
    """A local row that is not OK never replaces a remote OK row (another computer may have fetched it meanwhile)."""
    return [r for r in rows if r.get("status") == "OK" or (r["sr_no"], r["roll_no"]) not in ok_keys]


def sync_notification(client, results_xlsx, ledger=None, force=False):
    """Upload one notification (roll numbers + marksheets). Skipped when the files did not change since the last sync."""
    rx = Path(results_xlsx)
    mx = marksheet_path(rx)
    if ledger is not None and not force and ledger.synced(rx.stem, rx, mx):
        return {"skipped": True}
    year = rx.parent.name
    res = read_results(rx, year)
    out = {"skipped": False}
    r_rows = res.to_dict("records")
    for r in r_rows:
        r["year"] = int(year) if year.isdigit() else None
    out["results"] = upsert(client, "notification_results", r_rows, "notification,sr_no,roll_no")
    if mx.exists():
        ms = read_marksheets(mx, rx.stem, year)
        rows = drop_weaker(_ms_rows(ms, read_raw(mx)), remote_ok_keys(client, rx.stem))
        out["marksheets"] = upsert(client, "marksheets", rows, "notification,sr_no,roll_no")
    upsert(client, "notifications", [{
        "id": rx.stem, "year": int(year) if year.isdigit() else None, "title": rx.stem,
        "result_date": s(res["result_date"].iloc[0]) if len(res) else "", "rolls": int(len(res))}], "id")
    if ledger is not None:
        ledger.mark(rx.stem, rx, mx)
    return out


def sync_all(client, root, ledger=None, force=False):
    total = {"notifications": 0, "skipped": 0, "results": 0, "marksheets": 0}
    for fy in year_folders(root):
        for rf in results_files(fy):
            r = sync_notification(client, rf, ledger, force)
            if r.get("skipped"):
                total["skipped"] += 1
                continue
            total["notifications"] += 1
            total["results"] += r.get("results", 0)
            total["marksheets"] += r.get("marksheets", 0)
            print(f"   uploaded {rf.stem}: {r.get('results', 0)} roll numbers, {r.get('marksheets', 0)} marksheet rows")
        for f in fy.glob("exam_list_*.json"):
            try:
                push_exam_list(client, json.loads(f.read_text(encoding="utf-8")), f.stem.split("_")[-1])
            except Exception as e:
                print(f"   (exam list {f.name} not uploaded: {e})")
    if ledger is not None:
        ledger.save()
    return total


def push_exam_list(client, rows, year):
    data = [{"year": int(year) if str(year).isdigit() else None, "code": s(r.get("code")), "name": s(r.get("name")),
             "result_date": s(r.get("date"))} for r in rows if s(r.get("code"))]
    return upsert(client, "exam_list", data, "year,code,name")


def push_overrides(client, ov):
    rows = [{"key": k, "year": int(k.split("|")[0]) if k.split("|")[0].isdigit() else None, "course": v.get("course"),
             "exam_code": v.get("code"), "exam_name": v.get("name"), "skip": bool(v.get("skip")),
             "decided_at": v.get("decided_at"), "decided_by": v.get("by")} for k, v in ov.data.items()]
    return upsert(client, "exam_overrides", rows, "key") if rows else 0


def pull_overrides(client, ov):
    """Add the choices stored in Supabase that this PC does not have (or that are newer)."""
    n = 0
    for r in fetch_all(client, "exam_overrides", order=("key",)):
        cur = ov.get(r["key"])
        if cur is None or s(r.get("decided_at")) > s(cur.get("decided_at")):
            ov.data[r["key"]] = {"code": s(r.get("exam_code")), "name": s(r.get("exam_name")),
                                 "course": s(r.get("course")), "skip": bool(r.get("skip")),
                                 "decided_at": s(r.get("decided_at")), "by": s(r.get("decided_by"))}
            n += 1
    return n


_MS_SELECT = ",".join(c for c in MS_COLS if c not in ("year", "course", "result", "result_date"))


_MS_DB_COLS = [c for c in MS_COLS if c not in ("year", "course", "result", "result_date")]


def _assemble(res_rows, ms_rows):
    """Rows from notification_results and marksheets -> the 'combined' table (one row per roll number)."""
    res = pd.DataFrame(res_rows).reindex(columns=RES_COLS).fillna("")
    res["year"] = res["year"].map(lambda v: str(int(v)) if s(v).replace(".0", "").isdigit() else s(v))
    ms = pd.DataFrame(ms_rows).reindex(columns=_MS_DB_COLS)
    ms["fields"] = ms["fields"].map(lambda v: v if isinstance(v, dict) else {})
    for c in ("status", "note", "matched_exam", "exam_code", "fetched_at", "name", "father_name", "mother_name",
              "enrollment_no", "final_result", "final_total"):
        ms[c] = ms[c].fillna("")
    ms["match_score"] = pd.to_numeric(ms["match_score"], errors="coerce")
    full_ms = ms.merge(res[KEYS + ["year", "course", "result", "result_date"]], on=KEYS, how="left")
    return res, full_ms, combine(res, full_ms)


def load_supabase(client=None):
    """EVERYTHING in memory - only for small databases (the app uses the sb_* functions below instead)."""
    client = client or get_client()
    res_rows = fetch_all(client, "notification_results", order=KEYS)
    ms_rows = fetch_all(client, "marksheets", _MS_SELECT, order=KEYS)
    res, full_ms, comb = _assemble(res_rows, ms_rows)
    return {"results": res, "marksheets": full_ms, "combined": comb}


# ---- big databases: ask Supabase only for what a page shows ---------------------------------------------
def sb_course_summary(client=None):
    """One row per notification + course (view course_summary in supabase_views.sql)."""
    client = client or get_client()
    df = pd.DataFrame(fetch_all(client, "course_summary", order=("notification", "sr_no")))
    if df.empty:
        return df
    df["year"] = df["year"].map(lambda v: str(int(v)) if s(v).replace(".0", "").isdigit() else s(v))
    for c in ("statuses", "tried", "course", "result_date", "refreshed_at"):
        df[c] = df[c].fillna("") if c in df else ""
    return df


def refresh_summary(client=None):
    """Recount the course summary the app shows (SQL function refresh_course_summary)."""
    (client or get_client()).rpc("refresh_course_summary", {}).execute()


def sb_course(client, notification, sr_no):
    """All roll numbers of one course of one notification, with their marksheets."""
    client = client or get_client()
    res = fetch_all(client, "notification_results", order=("roll_no",), where={"notification": notification, "sr_no": sr_no})
    ms = fetch_all(client, "marksheets", _MS_SELECT, order=("roll_no",), where={"notification": notification, "sr_no": sr_no})
    return _assemble(res, ms)[2]


def _results_of(client, ms_rows, extra=()):
    keys = {(r["notification"], r["sr_no"], r["roll_no"]) for r in ms_rows}
    rolls = sorted({k[2] for k in keys})
    found = {}
    for i in range(0, len(rolls), 100):
        for r in (client.table("notification_results").select("*").in_("roll_no", rolls[i:i + 100]).execute().data or []):
            found[(r["notification"], r["sr_no"], r["roll_no"])] = r
    out = [r for k, r in found.items() if k in keys]
    seen = {(r["notification"], r["sr_no"], r["roll_no"]) for r in out}
    return out + [r for r in extra if (r["notification"], r["sr_no"], r["roll_no"]) not in seen]


def _clean_tok(t):
    return re.sub(r"[%_\\*(),]", "", t)


def sb_search(client, q, mode="Auto-detect", limit=600, summary=None):
    """Candidate rows (combined table) for a search; the caller ranks them with search(). Only the matching rows leave
    the database, so it works with hundreds of thousands of students."""
    client = client or get_client()
    q = q.strip()
    toks = [x for x in (_clean_tok(t) for t in re.split(r"[\s,.]+", q)) if x]
    if not toks:
        return _assemble([], [])[2]
    digits = q.replace(" ", "").isdigit()
    code_like = bool(re.fullmatch(r"(?=.*\d)(?=.*[A-Za-z])[A-Za-z0-9/_-]{5,}", q))
    m = mode if mode != "Auto-detect" else ("Roll number" if digits else ("Enrollment number" if code_like else "Name / part of name"))

    def ms_where(column, tokens, prefix=False):
        x = client.table("marksheets").select(_MS_SELECT)
        for t in tokens:
            x = x.ilike(column, f"%{t[:3] if prefix else t}%")
        return x.limit(limit).execute().data or []

    ms_rows, extra = [], []
    if m == "Roll number":
        roll = q.replace(" ", "")
        ms_rows = client.table("marksheets").select(_MS_SELECT).eq("roll_no", roll).limit(limit).execute().data or []
        extra = client.table("notification_results").select("*").eq("roll_no", roll).limit(limit).execute().data or []
    elif m == "Enrollment number":
        ms_rows = client.table("marksheets").select(_MS_SELECT).ilike("enrollment_no", q).limit(limit).execute().data or []
        if not ms_rows and mode == "Auto-detect":      # letters+digits that is really a roll number
            extra = client.table("notification_results").select("*").eq("roll_no", q).limit(limit).execute().data or []
    elif m == "Father's / mother's name":
        ms_rows = ms_where("father_name", toks) + ms_where("mother_name", toks)
    elif m == "Course / exam":
        if summary is not None and len(summary):
            hay = (summary["course"].astype(str) + " " + summary["notification"].astype(str)).str.upper()
            hit = summary[[all(t.upper() in h for t in toks) for h in hay]].head(5)
            frames = [sb_course(client, r.notification, r.sr_no) for r in hit.itertuples()]
            return pd.concat(frames, ignore_index=True) if frames else _assemble([], [])[2]
        return _assemble([], [])[2]
    else:                                              # name or part of it
        ms_rows = ms_where("name", toks)
        if not ms_rows:
            ms_rows = ms_where("father_name", toks) + ms_where("mother_name", toks)
        if not ms_rows:                                # small spelling mistakes: students sharing the first letters of ANY
            for tk in sorted(toks, key=len, reverse=True)[:3]:   # word (longest words first), ranked afterwards by search()
                if len(tk) >= 3:
                    ms_rows += ms_where("name", [tk], prefix=True)
    dedup = {(r["notification"], r["sr_no"], r["roll_no"]): r for r in ms_rows}
    ms_rows = list(dedup.values())
    return _assemble(_results_of(client, ms_rows, extra), ms_rows)[2]


def sb_history(client, roll, enrollment=""):
    """Every notification a student appears in (same roll number or enrollment number)."""
    client = client or get_client()
    ms = client.table("marksheets").select(_MS_SELECT).eq("roll_no", roll).limit(300).execute().data or []
    if enrollment:
        ms += client.table("marksheets").select(_MS_SELECT).ilike("enrollment_no", enrollment).limit(300).execute().data or []
    extra = client.table("notification_results").select("*").eq("roll_no", roll).limit(300).execute().data or []
    dedup = {(r["notification"], r["sr_no"], r["roll_no"]): r for r in ms}
    ms = list(dedup.values())
    return _assemble(_results_of(client, ms, extra), ms)[2]


def load_exam_list_supabase(client=None):
    client = client or get_client()
    rows = fetch_all(client, "exam_list", order=("year", "code", "name"))
    return _dedupe_exams([{"code": r["code"], "name": r["name"], "date": r.get("result_date")} for r in rows])


def load(source, root):
    """source: 'Local files' or 'Supabase'  ->  {'results', 'marksheets', 'combined'}"""
    return load_supabase() if source == "Supabase" else load_local(root)


def get_raw(source, root, notification, roll, sr_no="", client=None):
    """The full text of one marksheet (loaded only when asked for)."""
    if source == "Supabase":
        try:
            q = ((client or get_client()).table("marksheets").select("raw_text").eq("notification", notification)
                 .eq("roll_no", roll))
            if sr_no:
                q = q.eq("sr_no", sr_no)
            rows = q.limit(1).execute().data or []
            return s(rows[0].get("raw_text")) if rows else ""
        except Exception:                                   # the column was dropped to save space
            return ""
    for fy in year_folders(root):
        mp = fy / f"{notification}_MARKSHEETS.xlsx"
        if mp.exists():
            return read_raw(mp).get(roll, "")
    return ""


# --------------------------------------------------------------------------
# Search
# --------------------------------------------------------------------------
SEARCH_MODES = ["Auto-detect", "Name / part of name", "Roll number", "Enrollment number",
                "Father's / mother's name", "Course / exam"]


def _col(df, name):
    return df[name].fillna("").astype(str) if name in df else pd.Series("", index=df.index)


def _contains_all(series_norm, tokens):
    mask = pd.Series(True, index=series_norm.index)
    for t in tokens:
        mask &= series_norm.str.contains(re.escape(t), regex=True)
    return mask


def _fuzzy_name_mask(df, tokens, column):
    names = _col(df, column).map(norm_text)
    vocab = sorted({w for n in names.unique() for w in n.split()})
    mask = pd.Series(True, index=df.index)
    for t in tokens:
        near = get_close_matches(t, vocab, n=6, cutoff=0.75)
        if not near:
            return pd.Series(False, index=df.index)
        pat = r"\b(?:" + "|".join(re.escape(w) for w in near) + r")\b"
        mask &= names.str.contains(pat, regex=True)
    return mask


def _by_name(df, q, column):
    toks = norm_text(q).split()
    if not toks:
        return df.head(0), ""
    names = _col(df, column).map(norm_text)
    mask = _contains_all(names, toks)
    note = ""
    if not mask.any():
        mask = _fuzzy_name_mask(df, toks, column)
        note = "No exact match - showing similar spellings." if mask.any() else ""
    hit = df[mask].copy()
    if hit.empty:
        return hit, note
    qn = " ".join(toks)
    n = names[mask]
    hit["_rank"] = [0 if x == qn else 1 if x.startswith(toks[0]) else 2 for x in n]
    return hit.sort_values(["_rank", column]).drop(columns="_rank"), note


def _by_roll(df, q):
    q = s(q)
    r = _col(df, "roll_no")
    for mask in (r == q, r.str.startswith(q), r.str.contains(re.escape(q))):
        if mask.any():
            return df[mask]
    return df.head(0)


def _by_enrollment(df, q):
    q = s(q).upper()
    e = _col(df, "enrollment_no").str.upper()
    for mask in (e == q, e.str.contains(re.escape(q))):
        if mask.any():
            return df[mask]
    return df.head(0)


def search(df, query, mode="Auto-detect"):
    """-> (matching rows, message). Handles names in any word order, half names, small typos."""
    q = s(query)
    if df is None or df.empty or not q:
        return (df.head(0) if df is not None else pd.DataFrame()), ""
    msg = ""
    if mode == "Roll number":
        out = _by_roll(df, q)
    elif mode == "Enrollment number":
        out = _by_enrollment(df, q)
    elif mode == "Name / part of name":
        out, msg = _by_name(df, q, "name")
    elif mode == "Father's / mother's name":
        a, m1 = _by_name(df, q, "father_name")
        b, m2 = _by_name(df, q, "mother_name")
        out, msg = pd.concat([a, b]).drop_duplicates(subset=KEYS), m1 or m2
    elif mode == "Course / exam":
        toks = norm_text(q).split()
        names = (_col(df, "course") + " " + _col(df, "matched_exam")).map(norm_text)
        out = df[_contains_all(names, toks)]
    else:                                                    # auto-detect
        if re.fullmatch(r"\d{4,}", q):
            out = _by_roll(df, q)
            if out.empty:
                out = _by_enrollment(df, q)
        elif re.search(r"\d", q) and re.fullmatch(r"[A-Za-z0-9/\-_.]+", q):
            out = _by_enrollment(df, q)
            if out.empty:
                out = _by_roll(df, q)
        else:
            out, msg = _by_name(df, q, "name")
            if out.empty:
                a, _ = _by_name(df, q, "father_name")
                b, _ = _by_name(df, q, "mother_name")
                out = pd.concat([a, b]).drop_duplicates(subset=KEYS)
                msg = "Not found as a student's name - these match a father's / mother's name." if len(out) else ""
    return out, msg


def papers_from_fields(fields):
    """fields {'PAPER | Theory I': '45', ...} -> table (paper x component)."""
    rows = {}
    for k, v in (fields or {}).items():
        if " | " in k:
            paper, comp = k.rsplit(" | ", 1)
            rows.setdefault(paper, {})[comp] = v
    return pd.DataFrame.from_dict(rows, orient="index").rename_axis("paper").reset_index() if rows else pd.DataFrame()


def summarize(df):
    """Per notification + course: students, marksheets found, pass / fail counts."""
    if df is None or df.empty:
        return pd.DataFrame()
    d = df.copy()
    d["final_u"] = d["final_result"].fillna("").str.upper()
    d["notif_u"] = d["result"].fillna("").str.upper()
    g = d.groupby(["notification", "sr_no", "course"], sort=False)
    out = g.agg(students=("roll_no", "count"),
                marksheets_ok=("status", lambda x: int((x == "OK").sum())),
                not_found=("status", lambda x: int(x.isin(["NOT FOUND", "NO MATCH", "ERROR", "SKIPPED"]).sum())),
                not_fetched=("status", lambda x: int((x == "NOT FETCHED").sum())),
                pass_per_notification=("notif_u", lambda x: int((x == "PASS").sum())),
                fail_per_notification=("notif_u", lambda x: int((x == "FAIL").sum())),
                atkt_per_notification=("notif_u", lambda x: int(x.str.contains("ATKT|A.T.K.T").sum()))).reset_index()
    return out


# --------------------------------------------------------------------------
# Command line
# --------------------------------------------------------------------------
def main():
    args = sys.argv[1:]
    cmd = args[0].lower() if args else ""
    if cmd == "sync":
        client = get_client()
        led = Ledger(HERE / "ledger.json")
        ov = Overrides(HERE / "overrides.json")
        total = sync_all(client, HERE, led, force="--force" in args)
        n = push_overrides(client, ov)
        print(f"\nSupabase: {total}  (+{n} exam choices)")
    elif cmd == "counts":
        client = get_client()
        for t in ("notifications", "notification_results", "marksheets", "exam_overrides", "exam_list"):
            r = client.table(t).select("*", count="exact").limit(1).execute()
            print(f"{t:24} {r.count}")
    elif cmd == "status":
        led = Ledger(HERE / "ledger.json")
        for fy in year_folders(HERE):
            print(f"\n{fy.name}")
            print(plan(fy, led, need_sync=configured()).to_string(index=False))
    else:
        print(__doc__)


if __name__ == "__main__":
    main()
