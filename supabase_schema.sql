-- Run this once in Supabase: SQL Editor -> New query -> paste -> Run.
-- Then put SUPABASE_URL and SUPABASE_KEY (the service_role key; keep it private) in the .env file.
-- Safe to run again.

create extension if not exists pg_trgm;

-- one row per RESULT NOTIFICATION
create table if not exists notifications (
  id          text primary key,            -- the link name = file name without .pdf
  year        int,
  title       text,
  result_date text,
  rolls       int,
  updated_at  timestamptz default now()
);

-- every roll number of a notification, with the result printed there (PASS / FAIL / A.T.K.T.)
create table if not exists notification_results (
  notification text not null,
  sr_no        text not null,
  roll_no      text not null,
  year         int,
  course       text,
  result       text,
  result_date  text,
  primary key (notification, sr_no, roll_no)
);
create index if not exists nr_roll on notification_results (roll_no);

-- the marksheet of every roll number (fields = every other marksheet column, as JSON)
create table if not exists marksheets (
  notification  text not null,
  sr_no         text not null,
  roll_no       text not null,
  status        text,                      -- OK / NOT FOUND / NO MATCH / ERROR / SKIPPED
  note          text,
  matched_exam  text,
  exam_code     text,
  match_score   numeric,
  fetched_at    text,
  name          text,
  father_name   text,
  mother_name   text,
  enrollment_no text,
  final_result  text,
  final_total   text,
  obtained      numeric,
  maximum       numeric,
  percentage    numeric,
  fields        jsonb,
  raw_text      text,
  updated_at    timestamptz default now(),
  primary key (notification, sr_no, roll_no)
);
create index if not exists ms_roll      on marksheets (roll_no);
create index if not exists ms_enrol     on marksheets (upper(enrollment_no));
create index if not exists ms_name_trgm on marksheets using gin (upper(name) gin_trgm_ops);

-- the exam a human chose for a course (so it is not asked again)
create table if not exists exam_overrides (
  key        text primary key,             -- year|month|course words|semester|NEP
  year       int,
  course     text,
  exam_code  text,
  exam_name  text,
  skip       boolean default false,
  decided_at text,
  decided_by text
);

-- the site's list of declared exams (lets the app offer choices without a browser)
create table if not exists exam_list (
  year        int  not null,
  code        text not null,
  name        text not null,
  result_date text,
  primary key (year, code, name)
);

-- The scripts use the service_role key, which bypasses row level security.
-- Turn RLS on so the public (anon) key can read nothing.
alter table notifications        enable row level security;
alter table notification_results enable row level security;
alter table marksheets           enable row level security;
alter table exam_overrides       enable row level security;
alter table exam_list            enable row level security;
