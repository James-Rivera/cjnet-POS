# CJNET POS

A standalone cashier-first POS system for CJNET computer shop services: Xerox, printing, laminating, government-service assistance, custom items, sales records, expenses, reports, and editable prices.

Flowpress is not connected to this POS logic, data, or workflow. The shared design language is only a visual reference.

## Stack

- Next.js App Router
- TypeScript
- Tailwind CSS
- Supabase-ready data layer
- Vercel-ready project structure
- Local storage fallback for development without Supabase keys
- Supabase Auth login with protected routes when Supabase env vars are present

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
4. Add your public Supabase values:

```text
NEXT_PUBLIC_SUPABASE_URL=your-project-url
NEXT_PUBLIC_SUPABASE_ANON_KEY=your-anon-key
```

Do not commit real keys.

When Supabase env vars are present, the POS requires login and uses Supabase Auth plus RLS-protected tables. Without them, it uses browser local storage for quick local testing.

## Database Tables

- `profiles`
- `customers`
- `service_categories`
- `services`
- `price_settings`
- `sales`
- `sale_items`
- `expenses`

The schema supports sales with multiple line items, custom service prices, editable service prices, price history, expenses, date-range reports, daily summaries, and optional cashier/user profiles.

## POS Features

- Register with large clickable service cards
- Service search and category filters
- Custom-price services
- Discounts
- Cash received and change calculation
- Quick cash buttons
- Save sales with multiple sale items
- Toast confirmations for saved, edited, and deleted records
- Government-service references only for NBI, Police Clearance, PSA, and related online-service sales
- Sales history and CSV export
- Expense entry with CSV and Excel export
- Dashboard with today sales, net, expenses, transactions, best sellers, recent sales, and simple trend
- Monthly audit with gross sales, expenses, net income, average sale, top services, sales chart, and Excel export
- Reports by date range with Excel export
- Editable service prices
- Backup and restore JSON for local data
