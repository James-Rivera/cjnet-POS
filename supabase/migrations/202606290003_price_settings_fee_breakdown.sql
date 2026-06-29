-- Preserve owner-maintained fee breakdowns in price history.
-- Safe to rerun: only additive columns/constraints and non-destructive backfill.

alter table public.price_settings add column if not exists base_fee numeric(12, 2) not null default 0;
alter table public.price_settings add column if not exists service_fee numeric(12, 2) not null default 0;

do $$
begin
  if not exists (select 1 from pg_constraint where conname = 'price_settings_base_fee_nonnegative') then
    alter table public.price_settings add constraint price_settings_base_fee_nonnegative check (base_fee >= 0);
  end if;
  if not exists (select 1 from pg_constraint where conname = 'price_settings_service_fee_nonnegative') then
    alter table public.price_settings add constraint price_settings_service_fee_nonnegative check (service_fee >= 0);
  end if;
end $$;

-- Existing price history rows gain the current service split when the linked service has one.
-- This does not change sales, sale items, services, profiles, or audit logs.
update public.price_settings price_settings
set
  base_fee = services.base_fee,
  service_fee = services.service_fee
from public.services services
where price_settings.service_id = services.id
  and (price_settings.base_fee = 0 and price_settings.service_fee = 0)
  and (services.base_fee > 0 or services.service_fee > 0);
