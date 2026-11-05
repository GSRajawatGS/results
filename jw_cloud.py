"""
jw_cloud.py - many computers, one queue, no duplicated work.

Supabase is the manager. The PLANNER looks at which roll numbers still have no marksheet and puts them in the queue
(table fetch_jobs) in small batches. Any number of WORKERS - GitHub Actions runners, your own PCs, a cheap server -
repeat: take the next job from the queue (the database hands every job to exactly one worker), fetch those
marksheets, upload them, take the next one. A worker that dies simply loses its lease and the job goes to another.

    python jw_cloud.py plan [--retry] [--chunk 30] [--github-output]   queue the work that is still to do
    python jw_cloud.py work [--worker NAME] [--max-minutes 330] [--headed] [--no-stagger]
                                                                       fetch jobs until the queue is empty
    python jw_cloud.py status                                          what is queued / running / done / failed
    python jw_cloud.py reap                                            mark jobs lost too many times as failed

Setup: run supabase_schema.sql and supabase_queue.sql in Supabase, put SUPABASE_URL and SUPABASE_KEY in .env (or in
the GitHub repository secrets), and have result_scraper.py + result_jobs.py + marksheet_fetcher.py next to this file.
The exam a human picked for a course (Fix marksheets page) is shared through the exam_overrides table, so every
worker uses it; after a fix, run 'plan' again and only the affected roll numbers are queued.

Roll numbers are never fetched twice: the planner only queues roll numbers without a finished marksheet, the queue
refuses to hold the same batch twice, and uploads are upserts on (notification, sr_no, roll_no).
"""
import hashlib
import json
import os
import random
import shutil
import socket
import sys
import tempfile
import time
import traceback
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

import jw_store as js

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

CHUNK = int(os.environ.get("JW_CHUNK", "30"))                   # roll numbers per job
LEASE_SECONDS = int(os.environ.get("JW_LEASE_SECONDS", "900"))  # a job not finished within this is given to another worker
MAX_ATTEMPTS = 4
NO_MATCH_RETRY_HOURS = 12      # a course with no matching exam is looked at again only after this long


# --------------------------------------------------------------------------
# small helpers
# --------------------------------------------------------------------------
def _retry(fn, tries=4, what="request"):
    """Network hiccups must not lose a job: try again after 2, 4, 8 seconds."""
    for i in range(tries):
        try:
            return fn()
        except Exception as e:
            if i == tries - 1:
                raise
            wait = 2 ** (i + 1)
            print(f"   ({what} failed: {type(e).__name__}: {str(e)[:100]}; again in {wait} s)")
            time.sleep(wait)


def _recent(stamp, hours):
    try:
        at = datetime.strptime(js.s(stamp), "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return False
    return abs((datetime.now() - at).total_seconds()) < hours * 3600   # abs: PCs and runners are in different time zones


def needs_work(status, exam_code, fetched_at, ov_code, retry=False):
    """Does this roll number still need a marksheet fetch? Same rules as marksheet_fetcher.Store.done
    (finished = OK; NOT FOUND is looked at again only after a new exam choice or with retry), plus a cool-down so
    a course without a matching exam is not queued again and again."""
    st = js.s(status)
    if st == "OK":
        return False
    if st == "SKIPPED":
        return bool(ov_code) or retry
    if st == "NOT FOUND":
        return bool((ov_code and js.s(exam_code) != ov_code) or retry)
    if st == "NO MATCH" and not ov_code and not retry:
        return not _recent(fetched_at, NO_MATCH_RETRY_HOURS)
    return True                                  # never fetched, ERROR, ...


def _overrides(client):
    ov = js.Overrides(Path(tempfile.mkdtemp(prefix="jwov_")) / "overrides.json")
    try:
        js.pull_overrides(client, ov)
    except Exception as e:
        print(f"   (could not read the exam choices from Supabase: {type(e).__name__}: {str(e)[:100]})")
    return ov


def _year(v):
    t = js.s(v).replace(".0", "")
    return int(t) if t.isdigit() else None


# --------------------------------------------------------------------------
# the queue
# --------------------------------------------------------------------------
def reap(client):
    r = client.rpc("reap_fetch_jobs", {"p_max_attempts": MAX_ATTEMPTS}).execute().data
    return int(r or 0)


def claim(client, worker):
    rows = _retry(lambda: client.rpc("claim_fetch_job", {"p_worker": worker, "p_lease_seconds": LEASE_SECONDS,
                                                         "p_max_attempts": MAX_ATTEMPTS}).execute().data, what="claim")
    return rows[0] if rows else None


def finish(client, job, status, summary=None, error=None):
    """Only the worker that holds the job may finish it (so a job given to someone else is left alone)."""
    def go():
        client.table("fetch_jobs").update({
            "status": status, "summary": js.clean_json(summary), "last_error": error,
            "finished_at": f"{datetime.now(timezone.utc):%Y-%m-%dT%H:%M:%SZ}", "lease_until": None,
        }).eq("id", job["id"]).eq("worker", job["worker"]).execute()
    _retry(go, what="finish")


def _active(client):
    out = {}
    for st in ("pending", "running"):
        for r in _by_status(client, st):
            out[r["job_key"]] = st
    return out


def _by_status(client, status, columns="job_key", page=1000):
    rows, start = [], 0
    while True:
        r = (client.table("fetch_jobs").select(columns).eq("status", status).order("id")
             .range(start, start + page - 1).execute().data or [])
        rows += r
        if len(r) < page:
            return rows
        start += page


def plan(client, chunk=CHUNK, retry=False):
    """Queue every roll number that still needs a marksheet. Returns a summary dict."""
    import marksheet_fetcher as mf
    reaped = reap(client)
    out = {"reaped": reaped, "new_jobs": 0, "new_rolls": 0, "already_queued": 0, "up_to_date_rolls": 0, "skipped_courses": 0}
    notifs = [r["id"] for r in js.fetch_all(client, "notifications", "id", order=("id",))]
    if not notifs:
        print("   No roll numbers in Supabase yet - upload them first (jiwaji_results_scraper.py --central, or jw_store.py sync).")
        out["claimable"] = _claimable(client)
        return out
    ov = _overrides(client)
    active = _active(client)

    new_jobs = []
    for notif in notifs:                       # one notification at a time: small queries, no database timeouts
        res = pd.DataFrame(js.fetch_all(client, "notification_results", "notification,sr_no,roll_no,year,course,result,result_date",
                                        order=("sr_no", "roll_no"), where={"notification": notif}))
        if res.empty:
            continue
        known = {(r["sr_no"], r["roll_no"]): r for r in js.fetch_all(
            client, "marksheets", "sr_no,roll_no,status,exam_code,fetched_at", order=("sr_no", "roll_no"),
            where={"notification": notif})}
        for sr, g in res.groupby("sr_no", sort=False):
            course, rdate = js.s(g["course"].iloc[0]), js.s(g["result_date"].iloc[0])
            o = ov.get(mf.course_key(course, rdate))
            if o and o.get("skip"):
                out["skipped_courses"] += 1
                continue
            ov_code = o.get("code") if o and o.get("code") else None
            need = []
            for r in g.itertuples(index=False):
                m = known.get((r.sr_no, r.roll_no), {})
                if needs_work(m.get("status"), m.get("exam_code"), m.get("fetched_at"), ov_code, retry):
                    need.append((js.s(r.roll_no), js.s(r.result)))
                else:
                    out["up_to_date_rolls"] += 1
            need.sort()
            for i in range(0, len(need), chunk):
                part = need[i:i + chunk]
                key = f"{notif}|{sr}|{hashlib.sha1(','.join(r for r, _ in part).encode()).hexdigest()[:10]}"
                if key in active:
                    out["already_queued"] += 1
                    continue
                new_jobs.append({"job_key": key, "notification": notif, "sr_no": sr, "year": _year(g["year"].iloc[0]),
                                 "course": course, "result_date": rdate, "rolls": dict(part)})
                out["new_rolls"] += len(part)

    for i in range(0, len(new_jobs), 100):
        batch = new_jobs[i:i + 100]
        try:
            client.table("fetch_jobs").insert(batch).execute()
            out["new_jobs"] += len(batch)
        except Exception:                           # one of them was queued in the meantime: add the others one by one
            for j in batch:
                try:
                    client.table("fetch_jobs").insert(j).execute()
                    out["new_jobs"] += 1
                except Exception:
                    out["already_queued"] += 1
    out["claimable"] = _claimable(client)
    try:
        js.refresh_summary(client)                 # the app's course list shows the new counts
    except Exception as e:
        print(f"   (course summary not refreshed: {type(e).__name__}: {str(e)[:100]})")
    return out


def _claimable(client):
    return len(_by_status(client, "pending")) + len(_by_status(client, "running"))


# --------------------------------------------------------------------------
# a worker
# --------------------------------------------------------------------------
class SiteUnavailable(Exception):
    """The marksheet site cannot be used from this computer (blocked, down, or its page looks different)."""


MAX_SITE_FAILURES = 3          # in a row, then this worker stops instead of wasting time (and GitHub minutes)


def release(client, job, why):
    """Give the job back to the queue untouched (this attempt does not count)."""
    client.table("fetch_jobs").update({
        "status": "pending", "worker": None, "lease_until": None, "attempts": max(0, int(job["attempts"]) - 1),
        "last_error": why[:300]}).eq("id", job["id"]).eq("worker", job["worker"]).execute()


def run_job(client, p, br, rs, job, workdir, lists_pushed):
    """Fetch the marksheets of one job and upload them. Returns the counts dict."""
    import marksheet_fetcher as mf
    rolls = job["rolls"] if isinstance(job["rolls"], dict) else json.loads(job["rolls"])
    course, rdate, sr = js.s(job.get("course")), js.s(job.get("result_date")), js.s(job["sr_no"])
    year, _ = mf.course_year_session(course, rdate)
    cache_dir = workdir / str(year)                       # the site's exam list, read once per worker and year
    cache_dir.mkdir(parents=True, exist_ok=True)
    job_dir = workdir / f"job{job['id']}"
    job_dir.mkdir(parents=True, exist_ok=True)

    try:                                                   # without the site's exam list nothing can be matched:
        mf.exam_list(p, year, cache_dir)                   # stop here, write nothing, let the job go back to the queue
    except Exception as e:
        raise SiteUnavailable(f"exam list of {year}: {str(e)[:200]}")

    ov = js.Overrides(workdir / "overrides.json")         # the human's exam choices, always the latest
    try:
        js.pull_overrides(client, ov)
    except Exception as e:
        print(f"   (exam choices not refreshed: {type(e).__name__}: {str(e)[:100]})")

    g = pd.DataFrame({"Roll No": list(rolls), "Result": list(rolls.values()), "Result Date": rdate})
    store = mf.Store(job_dir / f"{job['notification']}_MARKSHEETS.xlsx")
    counts = {"OK": 0, "reused": 0, "NOT FOUND": 0, "ERROR": 0, "NO MATCH": 0, "SKIPPED": 0}
    mf.process_course(p, br, rs, store, ov, counts, cache_dir, sr, course, g, False, False, {})
    store.save()

    if store.path.exists():
        ms = js.read_marksheets(store.path, job["notification"], str(year))
        rows = js._ms_rows(ms, js.read_raw(store.path))
        rows = js.drop_weaker(rows, js.remote_ok_keys(client, job["notification"]))   # never replace a finished marksheet
        _retry(lambda: js.upsert(client, "marksheets", rows, "notification,sr_no,roll_no"), what="upload")
    lst = cache_dir / f"exam_list_{year}.json"
    if lst.exists() and year not in lists_pushed:         # the app's Fix page offers this list
        try:
            js.push_exam_list(client, json.loads(lst.read_text(encoding="utf-8")), year)
            lists_pushed.add(year)
        except Exception as e:
            print(f"   (exam list not uploaded: {type(e).__name__}: {str(e)[:100]})")
    shutil.rmtree(job_dir, ignore_errors=True)
    return counts


def work(client, worker, max_minutes=330, headless=True, stagger=True):
    """Take jobs from the queue until it is empty or max_minutes are over (a GitHub job may run 6 hours)."""
    import marksheet_fetcher as mf
    from playwright.sync_api import sync_playwright
    rs = mf._rs()
    rs.HEADLESS = headless                                # also used by the site's exam-list reader
    if stagger:                                           # many runners starting together should not hit the site together
        time.sleep(random.uniform(0, 15))
    deadline = time.time() + max_minutes * 60
    workdir = Path(tempfile.mkdtemp(prefix="jwworker_"))
    lists_pushed, total, site_fails = set(), {"jobs": 0, "failed": 0}, 0
    print(f"[{worker}] started; will stop after {max_minutes} minutes or when the queue is empty")
    with sync_playwright() as p:
        br = mf._Browser(p, headless)
        try:
            while time.time() < deadline:
                job = claim(client, worker)
                if not job:
                    print(f"[{worker}] no more jobs in the queue")
                    break
                n = len(job["rolls"]) if job.get("rolls") else 0
                print(f"\n[{worker}] job {job['id']}: {js.s(job.get('course'))[:70]}  ({n} roll numbers, attempt {job['attempts']})")
                t0 = time.time()
                try:
                    counts = run_job(client, p, br, rs, job, workdir, lists_pushed)
                    site_fails = 0
                    bad = counts.get("ERROR", 0)
                    finish(client, job, "failed" if bad else "done", {**counts, "seconds": int(time.time() - t0)},
                           f"{bad} roll number(s) had errors" if bad else None)
                    total["jobs"] += 1
                    total["failed"] += 1 if bad else 0
                    print(f"[{worker}] job {job['id']} {'had errors' if bad else 'done'}: {counts}")
                except KeyboardInterrupt:                 # Ctrl+C: give the job back at once instead of making it wait for its lease
                    try:
                        release(client, job, "stopped by the user")
                    finally:
                        raise
                except SiteUnavailable as e:
                    site_fails += 1
                    br.close()
                    release(client, job, f"SiteUnavailable: {e}")
                    print(f"[{worker}] the site is not usable from here ({e}); job {job['id']} given back "
                          f"({site_fails}/{MAX_SITE_FAILURES})")
                    if site_fails >= MAX_SITE_FAILURES:
                        print(f"[{worker}] STOPPING: the marksheet site does not work from this computer. "
                              f"Run the worker on a PC where the site opens in a browser.")
                        total["site_blocked"] = True
                        break
                    time.sleep(10)
                except Exception as e:                    # a bug or a broken page: report it and go on with the next job
                    traceback.print_exc()
                    br.close()
                    try:
                        finish(client, job, "failed", None, f"{type(e).__name__}: {str(e)[:300]}")
                    except Exception:
                        pass
                    total["failed"] += 1
        finally:
            br.close()
            for d in workdir.glob("*/debug"):              # screenshots of what the site showed (the workflow uploads them)
                try:
                    shutil.copytree(d, HERE / "debug_out" / d.parent.name, dirs_exist_ok=True)
                except Exception:
                    pass
            shutil.rmtree(workdir, ignore_errors=True)
    print(f"[{worker}] finished: {total}")
    if total.get("site_blocked"):
        sys.exit(3)                                       # shows as a failed worker in GitHub
    return total


# --------------------------------------------------------------------------
# status
# --------------------------------------------------------------------------
def status_frames(client):
    """(counts per status, jobs being worked on, failed jobs) as DataFrames."""
    rows = js.fetch_all(client, "fetch_jobs", "id,status,worker,attempts,lease_until,last_error,summary,rolls,course",
                        order=("id",))
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=["status", "jobs", "roll numbers"]), df, df
    df["n_rolls"] = df["rolls"].map(lambda r: len(r) if isinstance(r, dict) else 0)
    counts = df.groupby("status").agg(jobs=("id", "count"), roll_numbers=("n_rolls", "sum")).reset_index()
    counts = counts.rename(columns={"roll_numbers": "roll numbers"})
    running = df[df["status"] == "running"][["id", "worker", "attempts", "lease_until", "course"]]
    failed = df[df["status"] == "failed"][["id", "worker", "attempts", "last_error", "course"]]
    return counts, running, failed


def print_status(client):
    counts, running, failed = status_frames(client)
    if counts.empty:
        print("The queue is empty.")
        return
    print(counts.to_string(index=False))
    if len(running):
        print("\nbeing worked on:")
        print(running.to_string(index=False))
    if len(failed):
        print("\nfailed (the next 'plan' queues their roll numbers again):")
        print(failed.to_string(index=False))


# --------------------------------------------------------------------------
# command line
# --------------------------------------------------------------------------
def main():
    import argparse
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("command", choices=["plan", "work", "status", "reap", "refresh"])
    ap.add_argument("--retry", action="store_true", help="plan: also look again at NOT FOUND roll numbers")
    ap.add_argument("--chunk", type=int, default=CHUNK)
    ap.add_argument("--github-output", action="store_true", help="plan: write queued=<n> for the GitHub workflow")
    ap.add_argument("--worker", default=os.environ.get("JW_WORKER") or f"{socket.gethostname()}-{os.getpid()}")
    ap.add_argument("--max-minutes", type=float, default=330)
    ap.add_argument("--headed", action="store_true", help="work: show the browser (for your own PC)")
    ap.add_argument("--no-stagger", action="store_true")
    a = ap.parse_args()

    client = js.get_client()
    if a.command == "plan":
        r = plan(client, a.chunk, a.retry)
        print(f"Queued {r['new_jobs']} new job(s) with {r['new_rolls']} roll number(s). "
              f"Already queued: {r['already_queued']} job(s). Roll numbers up to date: {r['up_to_date_rolls']}. "
              f"Courses skipped by choice: {r['skipped_courses']}. Jobs waiting or running now: {r['claimable']}.")
        if a.github_output and os.environ.get("GITHUB_OUTPUT"):
            with open(os.environ["GITHUB_OUTPUT"], "a", encoding="utf-8") as f:
                f.write(f"queued={r['claimable']}\n")
    elif a.command == "work":
        work(client, a.worker, a.max_minutes, headless=not a.headed, stagger=not a.no_stagger)
    elif a.command == "refresh":
        js.refresh_summary(client)
        print("course summary refreshed")
    elif a.command == "reap":
        print(f"{reap(client)} lost job(s) marked failed")
    else:
        print_status(client)


if __name__ == "__main__":
    main()
