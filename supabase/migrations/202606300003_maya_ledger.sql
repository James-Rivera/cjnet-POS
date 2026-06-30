-- Adds Maya expected-balance tracking without touching existing sales, services, profiles, or audit records.
-- Safe to rerun: additive columns/tables/indexes/policies only.

alter table public.services add column if not exists uses_maya boolean not null default false;
alter table public.services add column if not exists maya_deduction_amount numeric(12, 2) not null default 0 check (maya_deduction_amount >= 0);
alter table public.services add column if not exists maya_deduction_mode text not null default 'pass_through' check (maya_deduction_mode in ('fixed', 'pass_through'));

create table if not exists public.maya_settings (
  id integer primary key default 1 check (id = 1),
  tracking_enabled boolean not null default false,
  current_balance numeric(12, 2) not null default 0 check (current_balance >= 0),
  low_balance_threshold numeric(12, 2) not null default 500 check (low_balance_threshold >= 0),
  updated_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.maya_ledger_entries (
  id uuid primary key default gen_random_uuid(),
  entry_type text not null check (entry_type in ('sale_deduction', 'top_up', 'adjustment')),
  amount numeric(12, 2) not null check (amount <> 0),
  direction text not null check (direction in ('in', 'out', 'adjustment')),
  balance_after numeric(12, 2),
  sale_id uuid references public.sales(id) on delete set null,
  sale_item_id uuid references public.sale_items(id) on delete set null,
  service_id text references public.services(id) on delete set null,
  notes text not null default '',
  reason text not null default '',
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now()
);

create index if not exists idx_maya_ledger_created_at on public.maya_ledger_entries(created_at);
create index if not exists idx_maya_ledger_entry_type on public.maya_ledger_entries(entry_type);
create index if not exists idx_maya_ledger_sale_id on public.maya_ledger_entries(sale_id);
create index if not exists idx_maya_ledger_service_id on public.maya_ledger_entries(service_id);
create index if not exists idx_services_uses_maya on public.services(uses_maya);

insert into public.maya_settings (id, tracking_enabled, current_balance, low_balance_threshold)
values (1, false, 0, 500)
on conflict (id) do nothing;

drop trigger if exists maya_settings_updated_at on public.maya_settings;
create trigger maya_settings_updated_at
before update on public.maya_settings
for each row execute function public.touch_updated_at();

alter table public.maya_settings enable row level security;
alter table public.maya_ledger_entries enable row level security;

drop policy if exists "Maya settings readable by active users" on public.maya_settings;
create policy "Maya settings readable by active users" on public.maya_settings
for select to authenticated
using (true);

drop policy if exists "Maya settings owner managed" on public.maya_settings;
create policy "Maya settings owner managed" on public.maya_settings
for all to authenticated
using (public.is_owner())
with check (public.is_owner());

drop policy if exists "Maya ledger readable by active users" on public.maya_ledger_entries;
create policy "Maya ledger readable by active users" on public.maya_ledger_entries
for select to authenticated
using (true);

drop policy if exists "Maya ledger sale deductions by cashiers" on public.maya_ledger_entries;
create policy "Maya ledger sale deductions by cashiers" on public.maya_ledger_entries
for insert to authenticated
with check (
  entry_type = 'sale_deduction'
  and direction = 'out'
  and created_by = auth.uid()
);

drop policy if exists "Maya ledger manual entries by owner manager" on public.maya_ledger_entries;
create policy "Maya ledger manual entries by owner manager" on public.maya_ledger_entries
for insert to authenticated
with check (
  public.is_owner_or_manager()
  and entry_type in ('top_up', 'adjustment')
  and created_by = auth.uid()
);

drop policy if exists "Maya ledger updates disabled" on public.maya_ledger_entries;
create policy "Maya ledger updates disabled" on public.maya_ledger_entries
for update to authenticated
using (false)
with check (false);

drop policy if exists "Maya ledger deletes owner only" on public.maya_ledger_entries;
create policy "Maya ledger deletes owner only" on public.maya_ledger_entries
for delete to authenticated
using (public.is_owner());

drop trigger if exists audit_maya_settings on public.maya_settings;
create trigger audit_maya_settings
after insert or update or delete on public.maya_settings
for each row execute function public.log_audit_event();

drop trigger if exists audit_maya_ledger_entries on public.maya_ledger_entries;
create trigger audit_maya_ledger_entries
after insert or update or delete on public.maya_ledger_entries
for each row execute function public.log_audit_event();
