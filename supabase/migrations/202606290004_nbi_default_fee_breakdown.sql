-- Make the default NBI service one-tap for cashiers while preserving owner changes.
-- Safe to rerun: only updates the known default NBI row when it has no existing fee split.

insert into public.service_categories (id, name, sort_order, is_active)
values ('online-services', 'Online Services', 30, true)
on conflict (id) do update
set name = excluded.name,
    sort_order = excluded.sort_order,
    is_active = true,
    updated_at = now();

update public.services
set
  name = 'NBI clearance',
  category_id = 'online-services',
  category = 'Online Services',
  option_label = 'NBI',
  price = 250,
  is_custom_price = false,
  group_name = 'NBI',
  base_fee = 160,
  service_fee = 90,
  is_active = true,
  updated_at = now()
where id = 'nbi-custom'
  and base_fee = 0
  and service_fee = 0
  and (
    (price = 0 and is_custom_price = true)
    or price = 250
  );

insert into public.services (
  id,
  name,
  category_id,
  category,
  option_label,
  price,
  is_custom_price,
  group_name,
  requires_tracking,
  base_fee,
  service_fee,
  sort_order,
  is_active
)
select
  'nbi-custom',
  'NBI clearance',
  'online-services',
  'Online Services',
  'NBI',
  250,
  false,
  'NBI',
  false,
  160,
  90,
  100,
  true
where not exists (
  select 1 from public.services where id = 'nbi-custom'
);

insert into public.price_settings (
  service_id,
  price,
  is_custom_price,
  base_fee,
  service_fee
)
select
  'nbi-custom',
  250,
  false,
  160,
  90
where exists (
  select 1
  from public.services
  where id = 'nbi-custom'
    and price = 250
    and base_fee = 160
    and service_fee = 90
)
on conflict (service_id) do update
set price = excluded.price,
    is_custom_price = excluded.is_custom_price,
    base_fee = excluded.base_fee,
    service_fee = excluded.service_fee
where price_settings.base_fee = 0
  and price_settings.service_fee = 0
  and (
    (price_settings.price = 0 and price_settings.is_custom_price = true)
    or price_settings.price = 250
  );
