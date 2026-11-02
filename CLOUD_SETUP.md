# Running the downloads on GitHub with many computers

Supabase is the manager. It holds a queue of small jobs (about 30 roll numbers each). Every worker takes the next job,
fetches the marksheets, uploads them, and takes another. The database gives each job to exactly one worker, so
nothing is fetched twice. A worker that dies loses its job after 15 minutes and another worker takes it.

## One-time setup
1. Supabase SQL editor: run `supabase_schema.sql`, then `supabase_queue.sql`.
2. Create a **private** GitHub repository (it handles student data - never make it public).
3. Put these files in it: `jiwaji_results_scraper.py`, `jw_store.py`, `jw_cloud.py`, `marksheet_fetcher.py`,
   `result_scraper.py`, `result_jobs.py` (your own files), `jw_app.py`, `requirements-cloud.txt`, `.gitignore`,
   `.github/workflows/jiwaji-fetch.yml`, and the two .sql files. Do not upload `.env`, PDFs or Excel files.
4. Repository -> Settings -> Secrets and variables -> Actions: add `SUPABASE_URL` and `SUPABASE_KEY` (service_role key).
5. Actions tab -> jiwaji-fetch -> Run workflow. Start with 2-4 workers, check the log, then raise it.
   It also runs every morning by itself.

## What happens in a run
- `plan` job: finds new RESULT NOTIFICATIONs, uploads their roll numbers, queues every roll number that has no marksheet.
- `work` jobs: N runners fetch from the queue for up to 5.5 hours, then stop. Run again for the rest.

## Also use your own PCs
On any PC with `.env` set: `python jw_cloud.py work --headed` (or without `--headed`). It joins the same queue.

## Commands
`python jw_cloud.py plan | work | status | reap`. The app's "What is done" page shows the queue and can queue work.

## Fixing a wrong exam
Use the app's Fix marksheets page (the choice is saved to Supabase and used by all workers), then queue the work again.
Only the affected roll numbers are fetched again.

## Cautions
- Untested against the live sites and real GitHub runners. The university sites may block or captcha GitHub's IPs; try 2 workers first.
- Logs hide student names (`JW_QUIET=1`). Keep it that way.
- Free plan: about 20 runners at once; private repos have a monthly minutes allowance - watch the usage page.
