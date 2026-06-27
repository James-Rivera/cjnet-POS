# CJNET POS UI Production Skill

You are working on the CJNET POS system, a real cashier-first web app for an internet cafe, Xerox, printing, and document service shop.

This is not a decorative dashboard project.
This is not a Dribbble concept.
This is not an AI-generated landing page.

The UI must feel like a real production POS:

* fast
* readable
* touch-friendly
* calm
* simple
* reliable
* cashier-focused
* owner-friendly
* not visually noisy

## Core Product Rule

Every UI decision must support one of these goals:

1. Help staff create a sale faster.
2. Help staff avoid mistakes.
3. Help the owner understand the business.
4. Help the system feel trustworthy and production-ready.

Do not add visual elements that do not support those goals.

## Visual Style Direction

CJNET uses a clean warm interface inspired by the existing Upload File page.

Use:

* soft beige app background
* white panels
* warm off-white surfaces
* rounded corners
* subtle shadows
* clear spacing
* calm typography
* yellow only for primary actions and selected states

Avoid:

* random gradients
* decorative background lines
* neon colors
* fake glassmorphism
* excessive shadows
* too many icons
* tiny unreadable text
* overdesigned cards
* AI-looking dashboard effects
* Dribbble-style visuals that hurt usability

## Design Principle

Do not make it “fancy.”
Make it feel real.

A good POS should look boring in the right ways:

* obvious actions
* clear totals
* large buttons
* strong receipt panel
* easy scanning
* no visual guessing

## Layout Rules

Desktop/tablet landscape POS layout:

* Left side: services/products
* Right side: current sale / receipt / checkout
* Top: navigation and daily summary
* Keep the current sale panel visible at all times on desktop/tablet landscape

Mobile/tablet portrait:

* Stack services and current sale
* Keep checkout actions easy to reach
* Do not create cramped two-column layouts on narrow screens

## Register Page Structure

The Register page must prioritize cashier workflow.

Recommended structure:

* Header/navigation
* Today summary strip
* Register title
* Main POS workspace

  * Services panel

    * section header
    * search
    * category filters
    * service grid
  * Current Sale panel

    * sale items
    * customer/reference
    * payment summary
    * cash input
    * quick cash buttons
    * change
    * save sale button

## Services Panel Rules

Service cards must be simple and scannable.

Each card should include:

1. service name
2. category/subtitle
3. price or custom price label

Do not use heavy typography inside cards.

Typography:

* service name: 13px or 14px, font-weight 600
* category/subtitle: 11px or 12px, font-weight 400, muted
* price badge: 11px or 12px, font-weight 700

Card rules:

* minimum width: 150px
* height: 104px to 120px
* border radius: 12px
* padding: 12px
* gap: 12px
* soft border only
* subtle hover state
* active/click state should feel tactile

Avoid:

* huge icons
* repeated icons on every card
* excessive bold text
* text wrapping into ugly multi-line blocks
* price badges that overpower service names
* cramped grids

## Current Sale / Receipt Panel Rules

The Current Sale panel is the most important part of the POS.

It should feel like a receipt and checkout area, not a normal form.

It must clearly show:

* item count
* selected items
* customer/reference
* subtotal
* discount
* total
* cash received
* change
* save sale action

Visual hierarchy:

* Total should be the strongest number
* Change should be clearly visible
* Save Sale should be the strongest action button
* Clear should be visible but not dominant

Empty state:
Use a calm empty state:

* small subtle icon or receipt symbol
* “No items yet”
* “Select a service to start a sale.”

Avoid:

* giant empty boxes with no guidance
* overly decorative empty states
* making Clear button visually stronger than Save Sale

## Dashboard Rules

The owner dashboard should be useful, not decorative.

Show:

* today’s sales
* today’s expenses
* net income
* transaction count
* best-selling services
* recent sales
* simple daily/weekly/monthly trend
* cashier activity if useful

Avoid:

* fake analytics widgets
* complicated graphs
* random circular charts
* decorative statistics that do not help the owner
* too many colors

Dashboard should answer:

* How much did we make today?
* How much did we spend?
* What is the net?
* What services sold most?
* Who made sales?
* What happened recently?

## Role-Based UI Rules

Owner navigation:

* Dashboard
* Register
* Sales
* Expenses
* Reports
* Prices
* Audit
* Staff
* Settings

Staff navigation:

* Register
* Expenses, only if enabled
* Sales, limited to own/today’s sales
* Upload/send file, if needed

Staff should not see:

* full dashboard
* net income
* long-term reports
* prices management
* audit logs
* staff management
* settings

Do not only hide routes in the UI.
Protected routes and database rules must enforce access.

## Typography Rules

Use a calm typography scale.

Recommended:

* page title: 24px, 700
* section title: 16px, 650 or 700
* card title: 13px or 14px, 600
* body text: 12px or 13px, 400 or 500
* labels: 11px or 12px, 500
* muted text: 11px or 12px, 400
* major amount/total: 22px to 28px, 700

Avoid:

* font-weight 800 or 900 except maybe for very large total numbers
* bolding every label
* using bold as decoration
* tiny text below 11px for important information

## Color Rules

Primary yellow should only be used for:

* primary buttons
* selected nav item
* selected category
* important highlight states

Do not use yellow everywhere.

Use neutral colors for most UI:

* background
* cards
* borders
* muted text
* secondary buttons

Danger/red only for:

* clear
* delete
* void
* negative net
* destructive actions

## Interaction Rules

Add clear but subtle interaction states:

* hover
* active/click
* selected
* disabled
* loading
* error

Service cards should feel clickable.
Buttons should have large enough touch targets.
Inputs should be easy to tap on tablets.

Minimum touch target:

* 40px height for normal controls
* 44px to 48px preferred for POS actions

## Spacing Rules

Use consistent spacing:

* 4px
* 8px
* 12px
* 16px
* 24px
* 32px

Do not use random spacing values unless necessary.

Main panels:

* 16px to 24px padding

Cards:

* 12px to 16px padding

Grid gaps:

* 12px to 16px

## Border and Shadow Rules

Use subtle borders and shadows.

Avoid:

* thick borders everywhere
* heavy shadows on every card
* multiple competing container styles

Recommended:

* main panel radius: 18px
* card/input/button radius: 12px
* border: soft warm gray
* shadow: subtle and consistent

## Component Rules

Create reusable components:

* AppShell
* Sidebar or TopNav
* SummaryCard
* ServiceCard
* CategoryPill
* ReceiptPanel
* PaymentSummary
* EmptyState
* PrimaryButton
* SecondaryButton
* DataTable
* DashboardCard

Do not create one-off styling for every page.

## Design Token Requirement

Use design tokens or Tailwind theme values for:

* colors
* spacing
* radius
* shadows
* typography

Do not hard-code random colors everywhere.

Suggested base palette:

* app background: #f4efe5
* card: #ffffff
* warm surface: #fff8e8
* primary yellow: #ffd400
* text: #151515
* muted text: #777777
* border: #eadfca
* danger: #e5484d
* success: #168a4a

## AI UI Anti-Pattern Checklist

Before finalizing any UI change, check:

* Did I add gradients just to make it look fancy?
* Did I add decorative lines that do not help usability?
* Did I make the interface harder to scan?
* Did I use too many font weights?
* Did I use too many colors?
* Did I make cards too small?
* Did I prioritize visuals over cashier speed?
* Did I create inconsistent spacing?
* Did I create components that do not match the rest of the system?

If yes, remove those changes.

## Final Quality Bar

The final UI should feel like:

* a real small-business POS
* clean enough for daily staff use
* professional enough for the owner
* simple enough for older tablets
* visually consistent with CJNET branding
* not obviously AI-generated

When in doubt:
Choose clarity over decoration.
Choose spacing over effects.
Choose hierarchy over gradients.
Choose usability over trendiness.
