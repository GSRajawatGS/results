-- Central job queue: many workers (GitHub Actions runners, your PCs ...) take jobs from here, never the same one twice.
-- Run this ONCE in the Supabase SQL editor, after supabase_schema.sql. Safe to run again.

-- One job = a small batch (about 30) of roll numbers of ONE course of ONE notification.
create table if not exists fetch_jobs (
  id            bigserial primary key,
  job_key       text not null,                       -- notification | sr_no | hash of the roll numbers
  notification  text not null,
  sr_no         text not null,
  year          int,
  course        text,
  result_date   text,
  rolls         jsonb not null,                      -- {"251001216": "PASS", ...}
  status        text not null default 'pending',     -- pending / running / done / failed
  attempts      int  not null default 0,
  worker        text,
  lease_until   timestamptz,                         -- a running job whose lease has run out is given to another worker
  created_at    timestamptz not null default now(),
  started_at    timestamptz,
  finished_at   timestamptz,
  summary       jsonb,                               -- {"OK": 28, "NOT FOUND": 2, ...}
  last_error    text
);

-- the same batch cannot be queued twice while it is waiting or running
create unique index if not exists fetch_jobs_active_uq on fetch_jobs (job_key) where status in ('pending', 'running');
create index if not exists fetch_jobs_status on fetch_jobs (status, id);

-- A worker calls this to get its next job. FOR UPDATE SKIP LOCKED makes it safe when many workers ask at the same
-- moment: every job goes to exactly one of them. A job whose worker disappeared (lease over) is handed out again,
-- at most p_max_attempts times.
create or replace function claim_fetch_job(p_worker text, p_lease_seconds int default 900, p_max_attempts int default 4)
returns setof fetch_jobs
language plpgsql
as $$
begin
  return query
  update fetch_jobs j
     set status      = 'running',
         worker      = p_worker,
         attempts    = j.attempts + 1,
         lease_until = now() + make_interval(secs => p_lease_seconds),
         started_at  = now()
   where j.id = (select f.id
                   from fetch_jobs f
                  where (f.status = 'pending' or (f.status = 'running' and f.lease_until < now()))
                    and f.attempts < p_max_attempts
                  order by f.id
                  for update skip locked
                  limit 1)
  returning j.*;
end;
$$;

-- Jobs that were lost too often are marked failed (the planner then queues their roll numbers again later).
create or replace function reap_fetch_jobs(p_max_attempts int default 4)
returns int
language plpgsql
as $$
declare n int;
begin
  update fetch_jobs
     set status = 'failed', finished_at = now(), lease_until = null,
         last_error = coalesce(last_error, 'worker lost and attempts used up')
   where status = 'running' and lease_until < now() and attempts >= p_max_attempts;
  get diagnostics n = row_count;
  return n;
end;
$$;

-- Only the service_role key (the scripts) may touch the queue: row level security on, no policies.
alter table fetch_jobs enable row level security;
revoke execute on function claim_fetch_job(text, int, int) from public, anon, authenticated;
revoke execute on function reap_fetch_jobs(int) from public, anon, authenticated;
