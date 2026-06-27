create extension if not exists pgcrypto;

create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  full_name text not null default '',
  role text not null default 'staff' check (role in ('owner', 'staff', 'manager', 'auditor')),
  status text not null default 'active' check (status in ('active', 'disabled')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.customers (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  reference text,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now()
);

create table if not exists public.service_categories (
  id text primary key,
  name text not null unique,
  sort_order integer not null default 100,
  is_active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.services (
  id text primary key,
  name text not null,
  category_id text references public.service_categories(id) on delete set null,
  category text not null,
  option_label text not null,
  price numeric(12, 2) not null default 0 check (price >= 0),
  is_custom_price boolean not null default false,
  is_active boolean not null default true,
  sort_order integer not null default 100,
  created_by uuid references auth.users(id) on delete set null,
  updated_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.services add column if not exists category_id text references public.service_categories(id) on delete set null;

create table if not exists public.price_settings (
  id uuid primary key default gen_random_uuid(),
  service_id text not null references public.services(id) on delete cascade,
  price numeric(12, 2) not null default 0 check (price >= 0),
  is_custom_price boolean not null default false,
  effective_from timestamptz not null default now(),
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now()
);

create table if not exists public.sales (
  id uuid primary key default gen_random_uuid(),
  sale_number bigint generated always as identity,
  customer_id uuid references public.customers(id) on delete set null,
  customer_note text,
  status text not null default 'completed' check (status in ('completed', 'voided')),
  subtotal numeric(12, 2) not null check (subtotal >= 0),
  discount numeric(12, 2) not null default 0 check (discount >= 0),
  total numeric(12, 2) not null check (total >= 0),
  cash_received numeric(12, 2) not null default 0 check (cash_received >= 0),
  change_due numeric(12, 2) not null default 0 check (change_due >= 0),
  cashier_id uuid references auth.users(id) on delete set null,
  voided_by uuid references auth.users(id) on delete set null,
  voided_at timestamptz,
  void_reason text,
  sold_at timestamptz not null default now(),
  created_at timestamptz not null default now()
);

create table if not exists public.sale_items (
  id uuid primary key default gen_random_uuid(),
  sale_id uuid not null references public.sales(id) on delete cascade,
  service_id text references public.services(id) on delete set null,
  service_name text not null,
  category text not null,
  option_label text not null,
  quantity integer not null check (quantity > 0),
  unit_price numeric(12, 2) not null check (unit_price >= 0),
  line_total numeric(12, 2) not null check (line_total >= 0),
  created_at timestamptz not null default now()
);

create table if not exists public.expenses (
  id uuid primary key default gen_random_uuid(),
  expense_date date not null default current_date,
  category text not null,
  description text,
  amount numeric(12, 2) not null check (amount > 0),
  created_by uuid references auth.users(id) on delete set null,
  updated_by uuid references auth.users(id) on delete set null,
  deleted_at timestamptz,
  created_at timestamptz not null default now()
);

create table if not exists public.settings (
  id integer primary key default 1 check (id = 1),
  staff_expenses_enabled boolean not null default false,
  updated_at timestamptz not null default now()
);

create table if not exists public.audit_logs (
  id uuid primary key default gen_random_uuid(),
  user_id uuid references auth.users(id) on delete set null,
  action text not null,
  table_name text not null,
  record_id text,
  old_value jsonb,
  new_value jsonb,
  created_at timestamptz not null default now()
);

create index if not exists idx_sales_sold_at on public.sales(sold_at);
create index if not exists idx_sale_items_sale_id on public.sale_items(sale_id);
create index if not exists idx_expenses_date on public.expenses(expense_date);
create index if not exists idx_services_category on public.services(category);
create index if not exists idx_services_category_id on public.services(category_id);
create index if not exists idx_price_settings_service_id on public.price_settings(service_id);
create unique index if not exists idx_price_settings_service_id_unique on public.price_settings(service_id);

create or replace function public.touch_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create or replace function public.is_owner()
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
      and role = 'owner'
      and status = 'active'
  );
$$;

create or replace function public.staff_expenses_enabled()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select coalesce((select staff_expenses_enabled from public.settings where id = 1), false);
$$;

create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, full_name, role, status)
  values (
    new.id,
    coalesce(new.raw_user_meta_data ->> 'full_name', new.raw_user_meta_data ->> 'name', ''),
    'staff',
    'active'
  )
  on conflict (id) do nothing;
  return new;
end;
$$;

create or replace function public.prevent_unauthorized_profile_updates()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  if not public.is_owner() then
    if new.role is distinct from old.role or new.status is distinct from old.status then
      raise exception 'Only owners can change role or status';
    end if;
  end if;
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
begin
  record_id := coalesce(new.id::text, old.id::text);

  insert into public.audit_logs (user_id, action, table_name, record_id, old_value, new_value)
  values (
    auth.uid(),
    lower(tg_op),
    tg_table_name,
    record_id,
    case when tg_op = 'DELETE' then to_jsonb(old) else to_jsonb(old) end,
    case when tg_op = 'DELETE' then null else to_jsonb(new) end
  );

  if tg_op = 'DELETE' then
    return old;
  end if;

  return new;
end;
$$;

alter table public.profiles enable row level security;
alter table public.customers enable row level security;
alter table public.service_categories enable row level security;
alter table public.services enable row level security;
alter table public.price_settings enable row level security;
alter table public.sales enable row level security;
alter table public.sale_items enable row level security;
alter table public.expenses enable row level security;
alter table public.settings enable row level security;
alter table public.audit_logs enable row level security;

drop policy if exists "Authenticated users can read profiles" on public.profiles;
drop policy if exists "Users can update their profile" on public.profiles;
drop policy if exists "Authenticated users can manage customers" on public.customers;
drop policy if exists "Authenticated users can manage service categories" on public.service_categories;
drop policy if exists "Authenticated users can manage services" on public.services;
drop policy if exists "Authenticated users can manage price settings" on public.price_settings;
drop policy if exists "Authenticated users can manage sales" on public.sales;
drop policy if exists "Authenticated users can manage sale items" on public.sale_items;
drop policy if exists "Authenticated users can manage expenses" on public.expenses;

create policy "Profiles are readable by owner or self" on public.profiles for select to authenticated using (auth.uid() = id or public.is_owner());
create policy "Profiles are insertable by owners" on public.profiles for insert to authenticated with check (public.is_owner());
create policy "Profiles are updateable by owner or self" on public.profiles for update to authenticated using (auth.uid() = id or public.is_owner()) with check (auth.uid() = id or public.is_owner());

create policy "Customers are owner managed" on public.customers for all to authenticated using (public.is_owner()) with check (public.is_owner());

create policy "Categories are readable by active users" on public.service_categories for select to authenticated using (is_active or public.is_owner());
create policy "Categories are owner managed" on public.service_categories for insert to authenticated with check (public.is_owner());
create policy "Categories can be updated by owner" on public.service_categories for update to authenticated using (public.is_owner()) with check (public.is_owner());
create policy "Categories can be deleted by owner" on public.service_categories for delete to authenticated using (public.is_owner());

create policy "Services are readable by active users" on public.services for select to authenticated using (is_active or public.is_owner());
create policy "Services are owner managed" on public.services for insert to authenticated with check (public.is_owner());
create policy "Services can be updated by owner" on public.services for update to authenticated using (public.is_owner()) with check (public.is_owner());
create policy "Services can be deleted by owner" on public.services for delete to authenticated using (public.is_owner());

create policy "Price settings are owner managed" on public.price_settings for all to authenticated using (public.is_owner()) with check (public.is_owner());

create policy "Sales are readable by owner or assigned cashier" on public.sales for select to authenticated using (public.is_owner() or cashier_id = auth.uid() or sold_at::date = current_date);
create policy "Sales are insertable by the active cashier" on public.sales for insert to authenticated with check (cashier_id = auth.uid() and status = 'completed');
create policy "Sales can be updated by owner" on public.sales for update to authenticated using (public.is_owner()) with check (public.is_owner());
create policy "Sales can be deleted by owner" on public.sales for delete to authenticated using (public.is_owner());

create policy "Sale items are readable with their sales" on public.sale_items for select to authenticated using (
  exists (
    select 1
    from public.sales
    where sales.id = sale_items.sale_id
      and (public.is_owner() or sales.cashier_id = auth.uid() or sales.sold_at::date = current_date)
  )
);
create policy "Sale items can be inserted with owned sales" on public.sale_items for insert to authenticated with check (
  exists (
    select 1
    from public.sales
    where sales.id = sale_items.sale_id
      and (public.is_owner() or sales.cashier_id = auth.uid())
  )
);
create policy "Sale items can be managed by owner" on public.sale_items for update to authenticated using (public.is_owner()) with check (public.is_owner());
create policy "Sale item deletes are owner only" on public.sale_items for delete to authenticated using (public.is_owner());

create policy "Expenses are readable by owner or the creating cashier" on public.expenses for select to authenticated using (public.is_owner() or created_by = auth.uid() or expense_date = current_date);
create policy "Expenses are insertable by owner or approved staff" on public.expenses for insert to authenticated with check (public.is_owner() or (public.staff_expenses_enabled() and created_by = auth.uid()));
create policy "Expenses are owner managed" on public.expenses for update to authenticated using (public.is_owner()) with check (public.is_owner());
create policy "Expense deletes are owner only" on public.expenses for delete to authenticated using (public.is_owner());

create policy "Settings are owner managed" on public.settings for all to authenticated using (public.is_owner()) with check (public.is_owner());

create policy "Audit logs are owner readable" on public.audit_logs for select to authenticated using (public.is_owner());

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
after insert on auth.users
for each row execute function public.handle_new_user();

drop trigger if exists profiles_updated_at on public.profiles;
create trigger profiles_updated_at
before update on public.profiles
for each row execute function public.touch_updated_at();

drop trigger if exists service_categories_updated_at on public.service_categories;
create trigger service_categories_updated_at
before update on public.service_categories
for each row execute function public.touch_updated_at();

drop trigger if exists services_updated_at on public.services;
create trigger services_updated_at
before update on public.services
for each row execute function public.touch_updated_at();

drop trigger if exists expenses_updated_at on public.expenses;
create trigger expenses_updated_at
before update on public.expenses
for each row execute function public.touch_updated_at();

drop trigger if exists settings_updated_at on public.settings;
create trigger settings_updated_at
before update on public.settings
for each row execute function public.touch_updated_at();

drop trigger if exists audit_service_categories on public.service_categories;
create trigger audit_service_categories
after insert or update or delete on public.service_categories
for each row execute function public.log_audit_event();

drop trigger if exists audit_services on public.services;
create trigger audit_services
after insert or update or delete on public.services
for each row execute function public.log_audit_event();

drop trigger if exists audit_price_settings on public.price_settings;
create trigger audit_price_settings
after insert or update or delete on public.price_settings
for each row execute function public.log_audit_event();

drop trigger if exists audit_sales on public.sales;
create trigger audit_sales
after insert or update or delete on public.sales
for each row execute function public.log_audit_event();

drop trigger if exists audit_sale_items on public.sale_items;
create trigger audit_sale_items
after insert or update or delete on public.sale_items
for each row execute function public.log_audit_event();

drop trigger if exists audit_expenses on public.expenses;
create trigger audit_expenses
after insert or update or delete on public.expenses
for each row execute function public.log_audit_event();

drop trigger if exists audit_profiles on public.profiles;
create trigger audit_profiles
after insert or update or delete on public.profiles
for each row execute function public.log_audit_event();

drop trigger if exists audit_settings on public.settings;
create trigger audit_settings
after insert or update or delete on public.settings
for each row execute function public.log_audit_event();

insert into public.settings (id, staff_expenses_enabled)
values (1, false)
on conflict (id) do update
set staff_expenses_enabled = excluded.staff_expenses_enabled,
    updated_at = now();

insert into public.service_categories (id, name, sort_order)
values
  ('xerox', 'Xerox', 10),
  ('printing', 'Printing', 20),
  ('online-services', 'Online Services', 30),
  ('finishing', 'Finishing', 40),
  ('custom', 'Custom', 50)
on conflict (id) do update
set name = excluded.name,
    sort_order = excluded.sort_order,
    is_active = true,
    updated_at = now();

insert into public.services (id, name, category_id, category, option_label, price, is_custom_price, sort_order)
values
  ('xerox-single', 'Xerox', 'xerox', 'Xerox', 'Single side', 3, false, 10),
  ('xerox-back', 'Xerox', 'xerox', 'Xerox', 'Back to back', 6, false, 20),
  ('print-bw-short', 'Black and white print', 'printing', 'Printing', 'Short', 5, false, 30),
  ('print-bw-long', 'Black and white print', 'printing', 'Printing', 'Long', 7, false, 40),
  ('print-bw-a4', 'Black and white print', 'printing', 'Printing', 'A4', 6, false, 50),
  ('print-color-short', 'Colored print', 'printing', 'Printing', 'Short', 10, false, 60),
  ('print-color-long', 'Colored print', 'printing', 'Printing', 'Long', 15, false, 70),
  ('print-color-a4', 'Colored print', 'printing', 'Printing', 'A4', 15, false, 80),
  ('gov-custom', 'Government service', 'online-services', 'Online Services', 'Custom price', 0, true, 90),
  ('nbi-custom', 'NBI assistance', 'online-services', 'Online Services', 'Custom price', 0, true, 100),
  ('police-clearance-custom', 'Police clearance', 'online-services', 'Online Services', 'Custom price', 0, true, 110),
  ('psa-custom', 'PSA assistance', 'online-services', 'Online Services', 'Custom price', 0, true, 120),
  ('laminating-custom', 'Laminating', 'finishing', 'Finishing', 'Custom price', 0, true, 130),
  ('misc-custom', 'Other shop service', 'custom', 'Custom', 'Custom price', 0, true, 140)
on conflict (id) do nothing;

update public.services
set category_id = case category
  when 'Xerox' then 'xerox'
  when 'Printing' then 'printing'
  when 'Online Services' then 'online-services'
  when 'Finishing' then 'finishing'
  else 'custom'
end
where category_id is null;

insert into public.price_settings (service_id, price, is_custom_price)
select services.id, services.price, services.is_custom_price
from public.services
where not exists (
  select 1
  from public.price_settings
  where price_settings.service_id = services.id
);
