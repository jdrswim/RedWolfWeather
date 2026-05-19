# HOA Manager — Self-managed HOA SaaS

A production-ready HOA management platform built with Next.js 14, Supabase, Tailwind CSS, and Stripe.

## Features

**Admin portal:**
- Dashboard with stats (owners, collected dues, outstanding balance, expenses)
- Owner management (add/view owners)
- Dues management (create individual or bulk dues, mark paid, filter by status)
- Expense tracking with category breakdown
- Document storage (upload/download bylaws, minutes, etc.)
- Messaging — direct messages to owners + broadcast announcements
- Settings — HOA configuration, payment rules, Stripe setup guide

**Owner portal:**
- Overview dashboard (outstanding balance, payment history)
- Dues list with one-click Stripe Checkout payment
- Messaging with HOA admin + view announcements
- Document library

**Platform features:**
- Email/password auth via Supabase Auth
- Role-based access (admin vs owner) enforced in middleware + RLS
- Realtime messaging via Supabase Realtime
- Stripe Checkout for dues payments
- Webhook updates payment & dues status
- Guided onboarding setup flow for new admins
- Setup checklist dashboard widget

---

## Quick Start

### 1. Clone and install

```bash
cd hoa-app
npm install
```

### 2. Set up Supabase

1. Create a project at [app.supabase.com](https://app.supabase.com)
2. Go to **SQL Editor** and run in order:
   - `supabase/schema.sql` — creates tables, triggers, indexes
   - `supabase/rls.sql` — enables RLS and creates all policies
   - `supabase/storage.sql` — creates the `hoa-documents` bucket

### 3. Set up Stripe

1. Create an account at [stripe.com](https://stripe.com)
2. Get your API keys from **Developers → API keys**
3. Create a webhook endpoint pointing to `https://yourdomain.com/api/stripe/webhook`
   - Listen for: `checkout.session.completed`, `checkout.session.expired`, `payment_intent.payment_failed`
4. Copy the webhook signing secret

> **Bank payouts** are configured separately in your Stripe Dashboard under  
> **Settings → Bank accounts and scheduling**. This app does not automatically  
> connect to your bank account — you must configure that in Stripe directly.

### 4. Configure environment variables

Copy `.env.example` to `.env.local` and fill in your values:

```bash
cp .env.example .env.local
```

```env
# Supabase
NEXT_PUBLIC_SUPABASE_URL=https://your-project.supabase.co
NEXT_PUBLIC_SUPABASE_ANON_KEY=eyJ...
SUPABASE_SERVICE_ROLE_KEY=eyJ...

# Stripe
STRIPE_SECRET_KEY=sk_test_...
NEXT_PUBLIC_STRIPE_PUBLIC_KEY=pk_test_...
STRIPE_WEBHOOK_SECRET=whsec_...

# App URL
NEXT_PUBLIC_APP_URL=http://localhost:3000
```

### 5. Run locally

```bash
npm run dev
```

Visit [http://localhost:3000](http://localhost:3000)

---

## Deployment (Vercel)

1. Push to GitHub
2. Import in [vercel.com](https://vercel.com)
3. Add all environment variables in Vercel project settings
4. Set `NEXT_PUBLIC_APP_URL` to your Vercel domain
5. Update your Stripe webhook URL to match your Vercel domain
6. Deploy

---

## First-time admin setup

1. Sign up at `/signup` and choose **HOA Admin / Board**
2. You'll be redirected to the **Setup Wizard** (`/onboarding`)
3. Complete all 5 steps:
   - **Organization** — HOA name, address, unit count
   - **Financial** — Stripe configuration notes + accounting start date
   - **Payment Rules** — due day, late fees, accepted methods
   - **Admin Profile** — confirm your display name
   - **Review & Launch** — verify settings and go live
4. The **Admin Dashboard** shows a Setup Checklist until all items are complete

---

## Database Schema

| Table | Description |
|-------|-------------|
| `profiles` | Extends `auth.users` — name, email, role, unit_number |
| `hoa_settings` | Single-row HOA config (name, dues rules, payment methods) |
| `dues` | Monthly dues per owner with status and balance |
| `payments` | Payment records linked to Stripe sessions |
| `messages` | Direct and broadcast messages with realtime |
| `expenses` | HOA operating expenses by category |
| `documents` | File metadata for Supabase Storage uploads |

---

## Architecture

```
app/
├── (auth)/          Login, signup pages
├── (dashboard)/
│   ├── admin/       Admin dashboard, owners, dues, expenses, docs, messages, settings
│   └── owner/       Owner overview, dues, messages, documents
├── onboarding/      5-step admin setup wizard
└── api/stripe/      Checkout session + webhook handler

lib/
├── supabaseClient.ts  Browser client
├── supabaseServer.ts  Server client + service role client
├── stripe.ts          Stripe SDK instance
└── utils.ts           Formatting helpers

middleware.ts          Route protection + role-based redirects
supabase/
├── schema.sql         Tables, triggers, indexes
├── rls.sql            Row Level Security policies
└── storage.sql        Storage bucket + policies
```

---

## Security

- All routes protected by `middleware.ts` — unauthenticated users redirected to `/login`
- Admin routes blocked for owner-role users
- Supabase RLS enforces data isolation at the database level — owners only see their own dues, payments, and messages
- Stripe webhook uses signature verification (`STRIPE_WEBHOOK_SECRET`)
- Stripe API keys stored as server-side environment variables only (never in the database)
