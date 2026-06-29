-- Adds a default GCash cash-in/cash-out service for pass-through amount + earned fee tracking.
-- Safe to rerun and does not alter existing transactions or profile records.

insert into public.service_categories (id, name, sort_order, is_active)
values ('gcash', 'GCash', 35, true)
on conflict (id) do update
set name = excluded.name,
    sort_order = excluded.sort_order,
    is_active = true,
    updated_at = now();

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
values (
  'gcash-cash-in',
  'GCash',
  'gcash',
  'GCash',
  'Cash in / Cash out',
  0,
  true,
  'GCash',
  false,
  0,
  0,
  130,
  true
)
on conflict (id) do update
set name = excluded.name,
    category_id = excluded.category_id,
    category = excluded.category,
    option_label = excluded.option_label,
    price = excluded.price,
    is_custom_price = excluded.is_custom_price,
    group_name = excluded.group_name,
    requires_tracking = excluded.requires_tracking,
    base_fee = excluded.base_fee,
    service_fee = excluded.service_fee,
    sort_order = excluded.sort_order,
    is_active = true,
    updated_at = now();
