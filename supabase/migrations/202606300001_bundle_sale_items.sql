-- Adds optional bundle/custom-reporting metadata to sale items.
-- Safe to rerun: additive columns only, no existing transactions are rewritten or deleted.

alter table public.sale_items add column if not exists bundle_id text;
alter table public.sale_items add column if not exists bundle_label text;
alter table public.sale_items add column if not exists is_uncategorized_custom boolean not null default false;

create index if not exists idx_sale_items_bundle_id on public.sale_items(bundle_id);
create index if not exists idx_sale_items_uncategorized_custom on public.sale_items(is_uncategorized_custom);
