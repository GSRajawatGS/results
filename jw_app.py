"""
jw_app.py - Streamlit app for the Jiwaji University results.

    pip install streamlit pandas openpyxl supabase python-dotenv pdfplumber selenium playwright
    streamlit run jw_app.py          (then open http://localhost:8501)

Keep it in the same folder as jw_store.py, marksheet_fetcher.py, jiwaji_results_scraper.py, result_scraper.py,
result_jobs.py. Pages:
    Search           name or part of a name, roll number, enrollment number, father's / mother's name, course
    Browse results   every notification and course: students, marksheets found, pass / fail counts
    Fix marksheets   roll numbers without a marksheet: choose the right exam by hand, then fetch again
    Run & sync       download new results, fetch marksheets, upload to Supabase (runs the scripts on this PC)
    What is done     which notification has its PDF / Excel / marksheets / upload done and what is left

Data source (sidebar): "Local files" (the Excel files next to the scripts) or "Supabase" (the cloud tables).
To protect the student data put  APP_PASSWORD=something  in the .env file: the app then asks for it.
"""
import os
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))

import jw_store as js                      # noqa: E402
import marksheet_fetcher as mf             # noqa: E402  (only its matching / key functions are used here)

LOGDIR = BASE / "app_logs"
SHOW_COLS = ["name", "father_name", "roll_no", "enrollment_no", "course", "result", "final_result", "final_total",
             "percentage", "status", "notification"]
UNRESOLVED = ["NO MATCH", "NOT FOUND", "ERROR", "SKIPPED"]


# ----------------------------------------------------------------------------
# helpers
# ----------------------------------------------------------------------------
class Runner:
    """Runs scripts one after the other in the background and keeps their output in a log file."""

    def __init__(self):
        self.thread, self.proc, self.title, self.log, self.stop_flag = None, None, "", None, False

    def running(self):
        return self.thread is not None and self.thread.is_alive()

    def start(self, cmds, title):
        LOGDIR.mkdir(exist_ok=True)
        self.title, self.stop_flag = title, False
        self.log = LOGDIR / f"{datetime.now():%Y%m%d_%H%M%S}.log"
        self.thread = threading.Thread(target=self._run, args=(cmds,), daemon=True)
        self.thread.start()

    def _run(self, cmds):
        env = dict(os.environ, PYTHONUNBUFFERED="1", PYTHONIOENCODING="utf-8")
        with open(self.log, "w", encoding="utf-8") as f:
            for cmd in cmds:
                if self.stop_flag:
                    break
                f.write("$ " + " ".join(str(c) for c in cmd) + "\n")
                f.flush()
                # no terminal here (stdin closed): the scripts never wait for typed answers
                self.proc = subprocess.Popen([str(c) for c in cmd], cwd=BASE, stdout=f, stderr=subprocess.STDOUT,
                                             stdin=subprocess.DEVNULL, env=env)
                self.proc.wait()
                if self.proc.returncode:
                    f.write(f"(exit code {self.proc.returncode})\n")

    def stop(self):
        self.stop_flag = True
        if self.proc and self.proc.poll() is None:
            self.proc.terminate()

    def tail(self, n=60):
        try:
            return "\n".join(Path(self.log).read_text(encoding="utf-8", errors="replace").splitlines()[-n:])
        except Exception:
            return ""


@st.cache_resource
def get_runner():
    return Runner()


def py(script, *args):
    return [sys.executable, str(BASE / script), *args]


def live_log(runner):
    st.divider()
    c1, c2 = st.columns([5, 1])
    c1.markdown(f"**Task:** {runner.title or '-'}   |   **Status:** {'running' if runner.running() else 'idle'}")
    if runner.running() and c2.button("Stop"):
        runner.stop()
        st.rerun()
    if runner.log:
        st.code(runner.tail(60) or "(no output yet)", language=None)
    if runner.running():
        time.sleep(2)
        st.rerun()


def csv_bytes(df):
    return df.to_csv(index=False).encode("utf-8-sig")        # opens correctly in Excel


@st.cache_data(ttl=600, show_spinner="Loading the data ...")
def load_data(source, root):
    if source == "Supabase":                      # big database: only the course list is loaded; pages ask for the rest
        return {"remote": True, "summary": js.sb_course_summary(), "combined": None}
    return js.load(source, root)


@st.cache_data(ttl=120, show_spinner="Searching ...")
def remote_search(q, mode):
    summary = load_data("Supabase", "")["summary"]
    return js.sb_search(js.get_client(), q, mode, summary=summary)


@st.cache_data(ttl=120)
def remote_course(notification, sr_no):
    return js.sb_course(js.get_client(), notification, sr_no)


@st.cache_data(ttl=120)
def remote_history(roll, enrollment):
    return js.sb_history(js.get_client(), roll, enrollment)


def refresh_button():
    """Supabase mode: the course list is a stored summary; this recounts it."""
    sm = load_data("Supabase", "")["summary"]
    when = sm["refreshed_at"].max() if len(sm) and "refreshed_at" in sm else ""
    c1, c2 = st.columns([4, 1])
    c1.caption(f"Course counts as of {str(when)[:16].replace('T', ' ')} UTC" if when else "Course counts: not calculated yet")
    if c2.button("Refresh the summary"):
        try:
            js.refresh_summary()
            load_data.clear()
            st.rerun()
        except Exception as e:
            st.error(f"{type(e).__name__}: {str(e)[:200]}  (run supabase_views.sql first)")


def need_data(source, root):
    try:
        data = load_data(source, root)
    except Exception as e:
        st.error(f"Could not load the data from {source}: {type(e).__name__}: {str(e)[:250]}")
        return None
    empty = data["summary"].empty if data.get("remote") else data["combined"].empty
    if empty:
        st.info("No data yet. " + ("Run 'Download' on the Run & sync page, then upload to Supabase."
                                   if source == "Supabase" else
                                   f"No notification Excel files were found under {root}."))
        return None
    return data


def exam_choices(source, root):
    if source == "Supabase":
        try:
            rows = js.load_exam_list_supabase()
            if rows:
                return rows
        except Exception:
            pass
    return js.local_exam_lists(root)


def table_of(df, cols=None):
    cols = [c for c in (cols or SHOW_COLS) if c in df]
    return df[cols].rename(columns=lambda c: c.replace("_", " "))


# ----------------------------------------------------------------------------
# Page: Search
# ----------------------------------------------------------------------------
def page_search(source, root):
    st.header("Search")
    data = need_data(source, root)
    if data is None:
        return
    remote = bool(data.get("remote"))
    df = None if remote else data["combined"]

    c1, c2 = st.columns([3, 2])
    q = c1.text_input("Search", placeholder="name, part of a name, roll number, enrollment number ...",
                      key="q").strip()
    mode = c2.selectbox("Search by", js.SEARCH_MODES, help="Auto-detect: digits = roll number, letters + digits = "
                                                            "enrollment number, letters = name (then parents' names).")
    base = data["summary"] if remote else df
    if remote:
        refresh_button()
    with st.expander("Filters"):
        f1, f2, f3 = st.columns(3)
        years = f1.multiselect("Year", sorted(base["year"].dropna().unique()))
        notifs = f2.multiselect("Notification", sorted(base["notification"].dropna().unique()))
        only_ok = f3.checkbox("Only students with a marksheet", value=False)

    if not q:
        st.info("Type a name (even half of it, words in any order, small spelling mistakes are tolerated), a roll "
                "number, or an enrollment number.")
        return
    if remote:
        try:
            df = remote_search(q, mode)
        except Exception as e:
            st.error(f"Search failed: {type(e).__name__}: {str(e)[:250]}  (did you run supabase_views.sql?)")
            return
    if years:
        df = df[df["year"].isin(years)]
    if notifs:
        df = df[df["notification"].isin(notifs)]
    if only_ok:
        df = df[df["status"] == "OK"]

    hits, msg = js.search(df, q, mode)
    if hits.empty:
        st.warning(f"Nothing found for '{q}'" + ("" if mode == "Auto-detect" else f" ({mode})") + ".")
        return
    if msg:
        st.info(msg)
    students = hits.drop_duplicates(subset=["roll_no", "name"]).shape[0]
    st.subheader(f"{len(hits)} result row(s)  -  {students} student(s)")
    shown = hits.head(500)
    if len(hits) > 500:
        st.caption("The first 500 rows are shown - type more of the name or add a filter.")
    st.dataframe(table_of(shown), hide_index=True)
    st.download_button("Download these rows (CSV)", csv_bytes(table_of(hits)), file_name="search_results.csv")

    st.subheader("One student")
    labels = [f"{r['name'] or '(no marksheet)'}  |  roll {r['roll_no']}  |  {str(r['course'])[:45]}  |  "
              f"{r['status']}  |  {r['notification'][:28]}" for _, r in shown.iterrows()]
    i = st.selectbox("Open", range(len(shown)), format_func=lambda k: labels[k])
    row = shown.iloc[i]
    a, b = st.columns(2)
    a.markdown("**Student**")
    info = {"Name": row["name"], "Father's name": row["father_name"], "Mother's name": row["mother_name"],
            "Roll number": row["roll_no"], "Enrollment number": row["enrollment_no"], "Course": row["course"],
            "Result in the notification": row["result"], "Result date": row["result_date"],
            "Marksheet result": row["final_result"], "Total": row["final_total"],
            "Percentage": row["percentage"], "Marksheet status": row["status"], "Exam": row["matched_exam"],
            "Exam code": row["exam_code"], "Note": row["note"], "Notification": row["notification"]}
    a.table(pd.DataFrame({"field": list(info), "value": [str(v) if str(v) not in ("None", "nan") else "" for v in info.values()]}))
    fields = row["fields"] if isinstance(row["fields"], dict) else {}
    b.markdown("**Other marksheet fields**")
    other = {k: v for k, v in fields.items() if " | " not in k}
    if other:
        b.table(pd.DataFrame({"field": list(other), "value": [str(v) for v in other.values()]}))
    else:
        b.info("No marksheet data for this row." if row["status"] != "OK" else "Nothing more on the marksheet.")
    papers = js.papers_from_fields(fields)
    if len(papers):
        st.markdown("**Papers**")
        st.dataframe(papers, hide_index=True)

    if remote:
        try:
            same = remote_history(row["roll_no"], row["enrollment_no"] or "")
        except Exception:
            same = df
    else:
        same = data["combined"]
    mask = same["roll_no"] == row["roll_no"]
    if row["enrollment_no"]:
        mask |= same["enrollment_no"] == row["enrollment_no"]
    history = same[mask]
    if len(history) > 1:
        st.markdown("**Every notification this student appears in** (same roll or enrollment number)")
        st.dataframe(table_of(history, ["roll_no", "course", "result", "final_result", "final_total", "percentage",
                                        "status", "notification"]), hide_index=True)
    with st.expander("Full marksheet text"):
        if row["status"] == "OK":
            try:
                txt = js.get_raw(source, root, row["notification"], row["roll_no"], row["sr_no"])
            except Exception as e:
                txt = f"(could not load: {type(e).__name__}: {str(e)[:150]})"
            st.code(txt or ("(the full text is not kept online, to save space - the fields above hold the marks)" if source == "Supabase" else "(not saved)"), language=None)
        else:
            st.caption("No marksheet for this row.")


# ----------------------------------------------------------------------------
# Page: Browse results
# ----------------------------------------------------------------------------
def page_browse(source, root):
    st.header("Browse results")
    data = need_data(source, root)
    if data is None:
        return
    remote = bool(data.get("remote"))
    if remote:
        refresh_button()
    df = data["summary"] if remote else data["combined"]
    years = sorted(df["year"].dropna().unique(), reverse=True)
    year = st.selectbox("Year", ["(all)"] + years)
    d = df if year == "(all)" else df[df["year"] == year]
    summ = d if remote else js.summarize(d)
    shown_cols = [c for c in summ.columns if c not in ("year", "result_date", "statuses", "tried")]
    st.dataframe(summ[shown_cols].rename(columns=lambda c: c.replace("_", " ")), hide_index=True)
    st.download_button("Download this summary (CSV)", csv_bytes(summ[shown_cols]), file_name="summary.csv")

    st.subheader("Students of one course")
    opts = [f"{r['notification'][:40]}  |  {r['sr_no']}. {r['course'][:70]}" for _, r in summ.iterrows()]
    if not opts:
        return
    k = st.selectbox("Notification and course", range(len(opts)), format_func=lambda i: opts[i])
    sel = summ.iloc[k]
    if remote:
        try:
            cls = remote_course(sel["notification"], sel["sr_no"])
        except Exception as e:
            st.error(f"Could not load this course: {type(e).__name__}: {str(e)[:250]}")
            return
    else:
        cls = d[(d["notification"] == sel["notification"]) & (d["sr_no"] == sel["sr_no"])]
    c = st.columns(4)
    c[0].metric("Students", len(cls))
    c[1].metric("Marksheets found", int((cls["status"] == "OK").sum()))
    c[2].metric("Without marksheet", int(cls["status"].isin(UNRESOLVED).sum()))
    pct = pd.to_numeric(cls["percentage"], errors="coerce")
    c[3].metric("Average %", round(float(pct.mean()), 2) if pct.notna().any() else "-")
    stat = st.multiselect("Status", sorted(cls["status"].unique()), default=sorted(cls["status"].unique()))
    table = cls[cls["status"].isin(stat)].sort_values("roll_no")
    st.dataframe(table_of(table), hide_index=True)
    st.download_button("Download this table (CSV)", csv_bytes(table_of(table)), file_name="course_students.csv")


# ----------------------------------------------------------------------------
# Page: Fix marksheets
# ----------------------------------------------------------------------------
def page_fix(source, root):
    st.header("Fix marksheets")
    st.caption("Roll numbers whose marksheet was not found - usually because the wrong exam link was picked. "
               "Choose the right exam once; it is remembered for that course and used from then on.")
    data = need_data(source, root)
    if data is None:
        return
    runner = get_runner()
    if data.get("remote"):
        sm = data["summary"]
        grp = (sm[sm["not_found"] > 0].rename(columns={"not_found": "without_marksheet", "statuses": "status"})
               [["notification", "sr_no", "course", "year", "without_marksheet", "status", "tried", "result_date"]]
               .reset_index(drop=True))
        bad = grp
    else:
        df = data["combined"]
        bad = df[df["status"].isin(UNRESOLVED)]
        grp = None if bad.empty else (bad.groupby(["notification", "sr_no", "course", "year"], sort=False)
               .agg(without_marksheet=("roll_no", "count"), status=("status", lambda x: ", ".join(sorted(set(x)))),
                    tried=("exam_code", lambda x: ", ".join(sorted({v for v in x if v}))),
                    result_date=("result_date", "first")).reset_index())
    if bad.empty:
        st.success("Every roll number has a marksheet.")
    else:
        st.dataframe(grp.drop(columns=["year"]).rename(columns=lambda c: c.replace("_", " ")), hide_index=True)
        labels = [f"{r['notification'][:36]}  |  {r['sr_no']}. {r['course'][:60]}  |  {r['without_marksheet']} roll(s)  |  "
                  f"{r['status']}" for _, r in grp.iterrows()]
        i = st.selectbox("Course to fix", range(len(grp)), format_func=lambda k: labels[k])
        g = grp.iloc[i]
        key = mf.course_key(g["course"], g["result_date"])
        ovs = js.Overrides(Path(root) / "overrides.json")      # the same file the scripts use
        cur = ovs.get(key)
        if cur:
            st.info("Your current choice: " + ("skip this course" if cur.get("skip") else f"{cur.get('code')}  -  {cur.get('name')}"))

        exams = exam_choices(source, root)
        if not exams:
            st.warning("The site's list of exams is not available yet. Run 'Refresh the site's exam list' on the "
                       "Run & sync page (and upload to Supabase if you use it).")
        else:
            words = st.text_input("Search the site's exam list (words of the name, or the exam code)", key=f"fx_{i}")
            ranked = mf.rank_exams(exams, g["course"], g["result_date"])
            if words.strip():
                ws = [w for w in words.upper().split() if w]
                ranked = [e for e in ranked if all(w in (str(e["name"]) + " " + str(e["code"])).upper() for w in ws)]
            ranked = ranked[:60]
            if not ranked:
                st.warning("No exam contains all those words.")
            else:
                pick = st.selectbox("Exam", range(len(ranked)),
                                    format_func=lambda k: f"{ranked[k]['code']}  |  {ranked[k]['score']:.2f}  |  "
                                                          f"{' '.join(str(ranked[k]['name']).split())[:90]}")
                c1, c2, c3 = st.columns(3)
                if c1.button("Use this exam", type="primary"):
                    ovs.set_exam(key, g["course"], str(ranked[pick]["code"]), str(ranked[pick]["name"]))
                    ovs.save()
                    _push_choice(ovs)
                    st.success("Saved. Now press 'Fetch the marksheets'.")
                if c2.button("Skip this course for good"):
                    ovs.set_skip(key, g["course"])
                    ovs.save()
                    _push_choice(ovs)
                    st.success("Saved: this course will not be fetched.")
                if cur and c3.button("Forget my choice"):
                    ovs.remove(key)
                    ovs.save()
                    st.success("Forgotten.")

        xlsx = Path(root) / str(g["year"]) / f"{g['notification']}.xlsx"
        st.divider()
        if data.get("remote"):
            st.caption("After choosing the exam, queue the work again; any computer running 'python jw_cloud.py work' "
                       "then fetches only these roll numbers.")
            if st.button("Queue these roll numbers again"):
                try:
                    import jw_cloud
                    r = jw_cloud.plan(js.get_client())
                    st.success(f"Queued {r['new_jobs']} job(s) with {r['new_rolls']} roll number(s).")
                except Exception as e:
                    st.error(f"{type(e).__name__}: {str(e)[:200]}")
        elif xlsx.exists():
            if st.button("Fetch the marksheets for this notification now (then upload to Supabase)",
                         disabled=runner.running()):
                cmds = [py("marksheet_fetcher.py", str(xlsx), "--no-ask")]
                if js.configured():
                    cmds.append(py("jw_store.py", "sync"))
                runner.start(cmds, f"fetch marksheets: {g['notification']}")
                st.rerun()
        else:
            st.caption(f"The Excel file {xlsx.name} is not on this PC, so the fetch has to be run on the PC that has "
                       f"the files (Run & sync page there).")

    with st.expander("All exam choices made so far"):
        ov_all = js.Overrides(Path(root) / "overrides.json")
        rows = [{"course": v.get("course"), "exam code": v.get("code"), "exam": v.get("name"),
                 "skip": v.get("skip"), "decided": v.get("decided_at"), "by": v.get("by")}
                for v in ov_all.data.values()]
        if rows:
            st.dataframe(pd.DataFrame(rows), hide_index=True)
        else:
            st.caption("None yet.")
    live_log(runner)


def _push_choice(ovs):
    if js.configured():
        try:
            js.push_overrides(js.get_client(), ovs)
        except Exception as e:
            st.warning(f"Saved here, but not uploaded to Supabase: {type(e).__name__}: {str(e)[:150]}")


# ----------------------------------------------------------------------------
# Page: Run & sync
# ----------------------------------------------------------------------------
def page_run(source, root):
    st.header("Run & sync")
    st.caption("These start the scripts on the PC where this app runs (they open a browser there). "
               "Nothing is repeated: finished notifications, roll numbers and uploads are skipped.")
    runner = get_runner()
    busy = runner.running()
    year = st.text_input("Year", str(datetime.now().year))
    s = py("jiwaji_results_scraper.py", year)

    st.subheader("Download and fetch")
    c1, c2, c3 = st.columns(3)
    if c1.button("Download new results + marksheets", disabled=busy):
        runner.start([[*s, "--no-ask"]], f"scraper {year}")
        st.rerun()
    if c2.button("Marksheets only (finish unfinished work)", disabled=busy):
        runner.start([[*s, "--marksheets-only", "--no-ask"]], f"marksheets {year}")
        st.rerun()
    if c3.button("Look again at NOT FOUND roll numbers", disabled=busy):
        runner.start([[*s, "--marksheets-only", "--retry", "--no-ask"]], f"retry {year}")
        st.rerun()
    st.caption("To be asked in the terminal which exam is right, run  python jiwaji_results_scraper.py " + year +
               "  yourself; or choose the exam on the 'Fix marksheets' page afterwards.")
    if st.button("Refresh the site's list of declared exams", disabled=busy):
        runner.start([py("marksheet_fetcher.py", "list", year)], f"exam list {year}")
        st.rerun()

    st.subheader("Supabase")
    if not js.configured():
        st.info("Supabase is not set up. 1) Run supabase_schema.sql in the Supabase SQL editor. 2) Put SUPABASE_URL and "
                "SUPABASE_KEY (the service_role key) in the .env file next to the scripts. 3) Restart the app.")
    force = st.checkbox("Upload everything again (even if unchanged)")
    d1, d2, d3 = st.columns(3)
    if d1.button("Upload local data to Supabase", disabled=busy or not js.configured()):
        runner.start([py("jw_store.py", "sync", *(["--force"] if force else []))], "upload to Supabase")
        st.rerun()
    if d2.button("Get exam choices from Supabase", disabled=not js.configured()):
        try:
            ov = js.Overrides(Path(root) / "overrides.json")
            n = js.pull_overrides(js.get_client(), ov)
            ov.save()
            st.success(f"{n} exam choice(s) copied from Supabase.")
        except Exception as e:
            st.error(f"{type(e).__name__}: {str(e)[:200]}")
    if d3.button("Reload the data in the app"):
        st.cache_data.clear()
        st.rerun()
    live_log(runner)


# ----------------------------------------------------------------------------
# Page: What is done
# ----------------------------------------------------------------------------
def page_status(source, root):
    st.header("What is done")
    led = js.Ledger(Path(root) / "ledger.json")
    folders = js.year_folders(root)
    if not folders:
        st.info(f"No year folders (like 2026) found under {root}.")
    for fy in folders:
        st.subheader(fy.name)
        t = js.plan(fy, led, need_sync=js.configured())
        if t.empty:
            st.caption("No PDFs downloaded yet.")
            continue
        st.dataframe(t.rename(columns={"excel": "roll numbers extracted"}), hide_index=True)
        open_ = int((t["to do"] != "-").sum())
        (st.warning if open_ else st.success)(f"{open_} notification(s) with work left" if open_ else "Everything is done.")

    st.subheader("Supabase")
    if not js.configured():
        st.caption("Not connected.")
        return
    if st.button("Count the rows in the Supabase tables"):
        try:
            client = js.get_client()
            rows = []
            for tname in ("notifications", "notification_results", "marksheets", "exam_overrides", "exam_list"):
                rows.append((tname, client.table(tname).select("*", count="exact").limit(1).execute().count))
            st.table(pd.DataFrame(rows, columns=["table", "rows"]))
        except Exception as e:
            st.error(f"{type(e).__name__}: {str(e)[:200]}")

    st.subheader("Work queue (many computers)")
    st.caption("Workers - GitHub Actions or any PC running 'python jw_cloud.py work' - take jobs from this queue, "
               "so no roll number is fetched twice.")
    try:
        import jw_cloud
        client = js.get_client()
        counts, running, failed = jw_cloud.status_frames(client)
        if counts.empty:
            st.caption("The queue is empty.")
        else:
            st.table(counts)
            if len(running):
                st.caption("Being worked on")
                st.dataframe(running, hide_index=True)
            if len(failed):
                st.caption("Failed (queue the work again to retry their roll numbers)")
                st.dataframe(failed, hide_index=True)
        c1, c2 = st.columns(2)
        if c1.button("Queue the work that is still to do"):
            r = jw_cloud.plan(client)
            st.success(f"Queued {r['new_jobs']} job(s) with {r['new_rolls']} roll number(s); "
                       f"{r['up_to_date_rolls']} roll numbers already have a marksheet.")
        if c2.button("Queue again, including NOT FOUND"):
            r = jw_cloud.plan(client, retry=True)
            st.success(f"Queued {r['new_jobs']} job(s) with {r['new_rolls']} roll number(s).")
    except Exception as e:
        st.warning(f"Queue not available: {type(e).__name__}: {str(e)[:200]} (did you run supabase_queue.sql?)")


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
def gate():
    js.load_env()
    secret = os.environ.get("APP_PASSWORD")
    if not secret or st.session_state.get("authed"):
        return True
    pw = st.text_input("Password", type="password")
    if pw and pw == secret:
        st.session_state["authed"] = True
        st.rerun()
    elif pw:
        st.error("Wrong password.")
    return False


def main():
    st.set_page_config(page_title="Jiwaji Results", page_icon="📊", layout="wide")
    try:                                         # Streamlit Cloud style secrets, if any
        for k in ("SUPABASE_URL", "SUPABASE_KEY", "APP_PASSWORD"):
            if k in st.secrets:
                os.environ.setdefault(k, st.secrets[k])
    except Exception:
        pass
    st.title("Jiwaji University - Results")
    if not gate():
        return
    with st.sidebar:
        page = st.radio("Page", ["Search", "Browse results", "Fix marksheets", "Run & sync", "What is done"])
        st.divider()
        sources = ["Local files", "Supabase"]
        source = st.radio("Load the data from", sources, index=1 if js.configured() else 0)
        if source == "Supabase" and not js.configured():
            st.error("Supabase is not configured (SUPABASE_URL / SUPABASE_KEY).")
        root = st.text_input("Folder with the year folders", str(BASE))
        if st.button("Reload the data"):
            st.cache_data.clear()
            st.rerun()
        st.caption(f"Source: {source}")
    root = root.strip() or str(BASE)
    {"Search": page_search, "Browse results": page_browse, "Fix marksheets": page_fix, "Run & sync": page_run,
     "What is done": page_status}[page](source, root)


if __name__ == "__main__":
    main()
