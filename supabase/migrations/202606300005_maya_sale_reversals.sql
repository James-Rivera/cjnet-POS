-- Allows Maya ledger to add money back when a Maya-backed sale is voided or hard-deleted.
-- The original sale_deduction rows remain in place for auditability; this only permits opposite reversal rows.

alter table public.maya_ledger_entries
  drop constraint if exists maya_ledger_entries_entry_type_check;

alter table public.maya_ledger_entries
  add constraint maya_ledger_entries_entry_type_check
  check (entry_type in ('sale_deduction', 'sale_reversal', 'top_up', 'adjustment'));

drop policy if exists "Maya ledger sale reversals by cashiers" on public.maya_ledger_entries;
create policy "Maya ledger sale reversals by cashiers" on public.maya_ledger_entries
for insert to authenticated
with check (
  entry_type = 'sale_reversal'
  and direction = 'in'
  and amount > 0
  and sale_id is not null
  and created_by = auth.uid()
  and exists (
    select 1
    from public.sales
    where sales.id = sale_id
      and (
        public.is_owner_or_manager()
        or (sales.cashier_id = auth.uid() and sales.status = 'voided' and sales.sold_at::date = current_date)
      )
  )
);
