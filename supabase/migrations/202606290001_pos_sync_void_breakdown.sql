-- Patch migration for POS sync, service grouping, fee breakdowns, and void workflows.
-- Safe to rerun: all schema additions are additive/idempotent and no production data is deleted.

-- Optional grouping/tracking/fee metadata for services. Existing services continue to work without these fields.
alter table public.services add column if not exists group_name text;
alter table public.services add column if not exists requires_tracking boolean not null default false;
alter table public.services add column if not exists base_fee numeric(12, 2) not null default 0;
alter table public.services add column if not exists service_fee numeric(12, 2) not null default 0;

do $$
begin
  if not exists (select 1 from pg_constraint where conname = 'services_base_fee_nonnegative') then
    alter table public.services add constraint services_base_fee_nonnegative check (base_fee >= 0);
  end if;
  if not exists (select 1 from pg_constraint where conname = 'services_service_fee_nonnegative') then
    alter table public.services add constraint services_service_fee_nonnegative check (service_fee >= 0);
  end if;
end $$;

-- Sales keep optional follow-up/reference intent and updated_at for audit-friendly mutation timestamps.
alter table public.sales add column if not exists needs_follow_up boolean not null default false;
alter table public.sales add column if not exists updated_at timestamptz not null default now();

-- Sale items preserve online assistance pass-through fee and actual CJNET revenue at the time of sale.
alter table public.sale_items add column if not exists base_fee numeric(12, 2) not null default 0;
alter table public.sale_items add column if not exists service_fee numeric(12, 2) not null default 0;
alter table public.sale_items add column if not exists pass_through_fee numeric(12, 2) not null default 0;
alter table public.sale_items add column if not exists revenue_amount numeric(12, 2) not null default 0;
alter table public.sale_items add column if not exists pricing_breakdown jsonb;

do $$
begin
  if not exists (select 1 from pg_constraint where conname = 'sale_items_base_fee_nonnegative') then
    alter table public.sale_items add constraint sale_items_base_fee_nonnegative check (base_fee >= 0);
  end if;
  if not exists (select 1 from pg_constraint where conname = 'sale_items_service_fee_nonnegative') then
    alter table public.sale_items add constraint sale_items_service_fee_nonnegative check (service_fee >= 0);
  end if;
  if not exists (select 1 from pg_constraint where conname = 'sale_items_pass_through_fee_nonnegative') then
    alter table public.sale_items add constraint sale_items_pass_through_fee_nonnegative check (pass_through_fee >= 0);
  end if;
  if not exists (select 1 from pg_constraint where conname = 'sale_items_revenue_amount_nonnegative') then
    alter table public.sale_items add constraint sale_items_revenue_amount_nonnegative check (revenue_amount >= 0);
  end if;
end $$;

-- Expenses gain soft-void metadata. Hard delete remains owner-only for exceptional cleanup.
alter table public.expenses add column if not exists updated_at timestamptz not null default now();
alter table public.expenses add column if not exists status text not null default 'active';
alter table public.expenses add column if not exists voided_by uuid references auth.users(id) on delete set null;
alter table public.expenses add column if not exists voided_at timestamptz;
alter table public.expenses add column if not exists void_reason text;

do $$
begin
  if not exists (select 1 from pg_constraint where conname = 'expenses_status_check') then
    alter table public.expenses add constraint expenses_status_check check (status in ('active', 'voided'));
  end if;
end $$;

create index if not exists idx_sales_status on public.sales(status);
create index if not exists idx_expenses_status on public.expenses(status);
create index if not exists idx_services_group_name on public.services(group_name);

-- Role helpers avoid changing profile records or existing profile policies.
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

-- Soft-void a sale. Staff can only void their own same-day transaction and must give a reason.
create or replace function public.void_sale(sale_id uuid, reason text)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  target_sale public.sales%rowtype;
begin
  if auth.uid() is null then
    raise exception 'Please sign in again.';
  end if;

  if nullif(trim(reason), '') is null then
    raise exception 'A void reason is required.';
  end if;

  select * into target_sale
  from public.sales
  where id = sale_id;

  if not found then
    raise exception 'Sale not found.';
  end if;

  if target_sale.status = 'voided' then
    return;
  end if;

  if public.is_owner_or_manager()
    or (target_sale.cashier_id = auth.uid() and target_sale.sold_at::date = current_date) then
    update public.sales
    set status = 'voided',
        voided_by = auth.uid(),
        voided_at = now(),
        void_reason = trim(reason),
        updated_at = now()
    where id = sale_id;
    return;
  end if;

  raise exception 'Only an owner, manager, or the same-day cashier can void this sale.';
end;
$$;

-- Soft-void an expense. Staff can only void their own same-day expense and must give a reason.
create or replace function public.void_expense(expense_id uuid, reason text)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  target_expense public.expenses%rowtype;
begin
  if auth.uid() is null then
    raise exception 'Please sign in again.';
  end if;

  if nullif(trim(reason), '') is null then
    raise exception 'A void reason is required.';
  end if;

  select * into target_expense
  from public.expenses
  where id = expense_id;

  if not found then
    raise exception 'Expense not found.';
  end if;

  if target_expense.status = 'voided' then
    return;
  end if;

  if public.is_owner_or_manager()
    or (target_expense.created_by = auth.uid() and target_expense.expense_date = current_date) then
    update public.expenses
    set status = 'voided',
        voided_by = auth.uid(),
        voided_at = now(),
        void_reason = trim(reason),
        updated_by = auth.uid(),
        updated_at = now()
    where id = expense_id;
    return;
  end if;

  raise exception 'Only an owner, manager, or the same-day creator can void this expense.';
end;
$$;

-- Updated-at triggers for tables now changed by the app/RPC functions.
drop trigger if exists sales_updated_at on public.sales;
create trigger sales_updated_at
before update on public.sales
for each row execute function public.touch_updated_at();

drop trigger if exists expenses_updated_at on public.expenses;
create trigger expenses_updated_at
before update on public.expenses
for each row execute function public.touch_updated_at();

-- Additive RLS policies for manager review and staff readback paths. Existing owner/staff policies remain intact.
drop policy if exists "Managers can read sales" on public.sales;
create policy "Managers can read sales" on public.sales for select to authenticated using (public.is_manager());

drop policy if exists "Managers can read sale items" on public.sale_items;
create policy "Managers can read sale items" on public.sale_items for select to authenticated using (public.is_manager());

drop policy if exists "Managers can read expenses" on public.expenses;
create policy "Managers can read expenses" on public.expenses for select to authenticated using (public.is_manager());

drop policy if exists "Managers can add expenses" on public.expenses;
create policy "Managers can add expenses" on public.expenses for insert to authenticated with check (public.is_manager() and created_by = auth.uid());

drop policy if exists "Managers can read audit logs" on public.audit_logs;
create policy "Managers can read audit logs" on public.audit_logs for select to authenticated using (public.is_manager());
