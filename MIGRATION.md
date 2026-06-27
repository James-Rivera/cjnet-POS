# CJNET POS Next.js Migration

This folder is the Vercel-ready migration of the original static POS into a standalone CJNET cashier system.

Flowpress is only a visual reference for the clean beige/white/yellow design language. POS data, sales, expenses, reports, prices, and cashier workflow are independent from Flowpress logic.

## Stack

- Next.js App Router
- TypeScript
- Tailwind CSS
- Supabase-ready database adapter

## Local Development

```powershell
npm install
npm run dev -- --port 4180
```

Open:

```text
http://localhost:4180
```

## Supabase Setup

1. Create a Supabase project.
2. Run `supabase/schema.sql` in the Supabase SQL editor.
3. Copy `.env.example` to `.env.local`.
4. Add:

```text
NEXT_PUBLIC_SUPABASE_URL=your-project-url
NEXT_PUBLIC_SUPABASE_ANON_KEY=your-anon-key
```

Do not commit real keys.

When Supabase env vars are present, the POS uses Supabase. Without them, it uses browser local storage for development and testing.

## Database Tables

- `profiles`
- `customers`
- `service_categories`
- `services`
- `price_settings`
- `sales`
- `sale_items`
- `expenses`

The schema supports sales with multiple line items, editable service prices, price history, expenses, date-range reports, daily summaries, and cashier/user profiles.

## Verified

- Register/cart calculations
- Custom services
- Discounts
- Cash/change
- Sales save and listing
- Expenses save and listing
- Reports: gross, expenses, net, transactions
- Excel exports for expenses, date-range reports, and monthly audit
- Monthly audit with chart and service/expense summaries
- Toast confirmations after write actions
- Government-service reference field only when needed
- Price editing persistence
- Backup JSON export
- Production build passes
