-- Adds end-of-day closing records for cashier money reconciliation.
-- Safe to rerun: additive table/index/policies only; no sales, expenses, services, profiles, or audit logs are rewritten.

create table if not exists public.daily_closings (
  id uuid primary key default gen_random_uuid(),
  closing_date date not null default current_date,
  opening_cash numeric(12, 2) not null default 0 check (opening_cash >= 0),
  expected_cash numeric(12, 2) not null default 0 check (expected_cash >= 0),
  actual_cash numeric(12, 2) not null default 0 check (actual_cash >= 0),
  cash_difference numeric(12, 2) not null default 0,
  wallet_balance numeric(12, 2) check (wallet_balance is null or wallet_balance >= 0),
  notes text not null default '',
  summary jsonb not null default '{}'::jsonb,
  closed_by uuid references auth.users(id) on delete set null,
  closed_at timestamptz not null default now(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (closing_date, closed_by)
);

create index if not exists idx_daily_closings_date on public.daily_closings(closing_date);
create index if not exists idx_daily_closings_closed_by on public.daily_closings(closed_by);

create or replace function public.touch_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create or replace function public.log_audit_event()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  record_id text;
  old_payload jsonb;
  new_payload jsonb;
begin
  if tg_op = 'INSERT' then
    record_id := new.id::text;
    old_payload := null;
    new_payload := to_jsonb(new);
  elsif tg_op = 'UPDATE' then
    record_id := new.id::text;
    old_payload := to_jsonb(old);
    new_payload := to_jsonb(new);
  elsif tg_op = 'DELETE' then
    record_id := old.id::text;
    old_payload := to_jsonb(old);
    new_payload := null;
  end if;

  insert into public.audit_logs (user_id, action, table_name, record_id, old_value, new_value)
  values (auth.uid(), lower(tg_op), tg_table_name, record_id, old_payload, new_payload);

  if tg_op = 'DELETE' then
    return old;
  end if;

  return new;
end;
$$;

create or replace function public.is_manager()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select exists (
    select 1
    from public.profiles
    where id = auth.uid()
      and role = 'manager'
      and status = 'active'
  );
$$;

create or replace function public.is_owner_or_manager()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select public.is_owner() or public.is_manager();
$$;

drop trigger if exists daily_closings_updated_at on public.daily_closings;
create trigger daily_closings_updated_at
before update on public.daily_closings
for each row execute function public.touch_updated_at();

alter table public.daily_closings enable row level security;

drop policy if exists "Daily closings are readable by owner manager or cashier" on public.daily_closings;
create policy "Daily closings are readable by owner manager or cashier" on public.daily_closings
for select to authenticated
using (public.is_owner_or_manager() or closed_by = auth.uid() or closing_date = current_date);

drop policy if exists "Cashiers can submit own daily closing" on public.daily_closings;
create policy "Cashiers can submit own daily closing" on public.daily_closings
for insert to authenticated
with check (closed_by = auth.uid());

drop policy if exists "Cashiers can update own daily closing" on public.daily_closings;
create policy "Cashiers can update own daily closing" on public.daily_closings
for update to authenticated
using (public.is_owner_or_manager() or closed_by = auth.uid())
with check (public.is_owner_or_manager() or closed_by = auth.uid());

drop policy if exists "Daily closing deletes are owner only" on public.daily_closings;
create policy "Daily closing deletes are owner only" on public.daily_closings
for delete to authenticated
using (public.is_owner());

drop trigger if exists audit_daily_closings on public.daily_closings;
create trigger audit_daily_closings
after insert or update or delete on public.daily_closings
for each row execute function public.log_audit_event();
