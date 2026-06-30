-- Makes end-of-day PEDIcash/drawer counting optional without rewriting closing summaries.
-- Existing closing records are treated as counted because they were created before the optional flow existed.

alter table public.daily_closings
  add column if not exists cash_counted boolean not null default true;

comment on column public.daily_closings.cash_counted is
  'True when the cashier counted physical PEDIcash/drawer cash. False when the closing was saved as summary-only.';
