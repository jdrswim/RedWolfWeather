# Kevin's HOA Management System

A production-ready HOA Management SaaS web app built with Next.js 14, Supabase, and Stripe.
Deployed at `redwolfweather.com/kevin`.

## Features

**Admin**
- Dashboard with financial summary and setup checklist
- Manage owners (add, edit, deactivate)
- Set and track dues per owner
- Record expenses (insurance, utilities, etc.)
- Send broadcast and direct messages
- Upload HOA documents (bylaws, rules)

**Owner**
- View dues balance and payment history
- Pay dues via Stripe Checkout
- Receive and send messages
- Download HOA documents

## Tech Stack

- **Frontend**: Next.js 14 (App Router, TypeScript)
- **Backend/Auth**: Supabase (Auth, Postgres, Storage, RLS, Realtime)
- **Styling**: TailwindCSS
- **Payments**: Stripe Checkout
- **Deployment**: Vercel + GitHub Pages (redwolfweather.com)

---

## Setup Guide

### 1. Supabase Setup

1. Go to [supabase.com](https://supabase.com) and create a new project
2. Go to **SQL Editor** and run the full contents of `supabase/schema.sql`
3. Go to **Storage** and create two buckets:
   - `hoa-documents` (set to private)
   - `receipts` (set to private)
4. Go to **Project Settings > API** and copy:
   - `Project URL` → `NEXT_PUBLIC_SUPABASE_URL`
   - `anon/public` key → `NEXT_PUBLIC_SUPABASE_ANON_KEY`
   - `service_role` key → `SUPABASE_SERVICE_ROLE_KEY` *(keep secret!)*
5. Go to **Authentication > URL Configuration** and set:
   - Site URL: `https://redwolfweather.com/kevin`
   - Redirect URLs: `https://redwolfweather.com/kevin/**`

### 2. Stripe Setup

> **Important**: Stripe keys must be entered as environment variables. The app does NOT store them in the database.

1. Go to [stripe.com](https://stripe.com) and create/log in to your account
2. Go to **Developers > API Keys** and copy:
   - Publishable key → `NEXT_PUBLIC_STRIPE_PUBLIC_KEY`
   - Secret key → `STRIPE_SECRET_KEY`
3. Set up Webhook:
   - Go to **Developers > Webhooks > Add endpoint**
   - URL: `https://redwolfweather.com/kevin/api/stripe/webhook`
   - Events: `checkout.session.completed`, `payment_intent.payment_failed`
   - Copy the signing secret → `STRIPE_WEBHOOK_SECRET`
4. **Bank Account**: Connect your bank in Stripe Dashboard under **Settings > Bank Accounts** (this is done directly in Stripe — not in the app)

### 3. Environment Variables

Copy `.env.example` to `.env.local` and fill in all values:

```bash
cp .env.example .env.local
```

```
NEXT_PUBLIC_SUPABASE_URL=https://xxxx.supabase.co
NEXT_PUBLIC_SUPABASE_ANON_KEY=eyJhbGci...
SUPABASE_SERVICE_ROLE_KEY=eyJhbGci...
STRIPE_SECRET_KEY=sk_live_...
NEXT_PUBLIC_STRIPE_PUBLIC_KEY=pk_live_...
STRIPE_WEBHOOK_SECRET=whsec_...
NEXT_PUBLIC_BASE_URL=https://redwolfweather.com
```

### 4. Local Development

```bash
cd kevin
npm install
npm run dev
# App runs at http://localhost:3000/kevin
```

### 5. Vercel Deployment

1. Push this repo to GitHub
2. Go to [vercel.com](https://vercel.com), create a new project
3. Set **Root Directory** to `kevin`
4. Add all environment variables from `.env.local`
5. Set custom domain to `redwolfweather.com` in Vercel
6. In your DNS settings, point `redwolfweather.com` to Vercel

### 6. First-Time Admin Onboarding

1. Navigate to `https://redwolfweather.com/kevin`
2. Click **Sign Up** and create your admin account
3. In Supabase SQL Editor, run:
   ```sql
   UPDATE profiles SET role = 'admin' WHERE email = 'your-admin@email.com';
   ```
4. Log back in — you'll be redirected to the Admin dashboard
5. Complete the **Setup Checklist** on the admin dashboard:
   - Enter HOA name, address, and number of units
   - Configure payment settings
   - Upload initial documents (bylaws, rules)

---

## Database Schema

| Table | Description |
|-------|-------------|
| `profiles` | User accounts (owners + admins) |
| `dues` | Monthly dues records per owner |
| `payments` | Stripe payment records |
| `messages` | Owner-admin messaging |
| `expenses` | HOA operating expenses |
| `documents` | Uploaded HOA documents |
| `hoa_settings` | HOA configuration |
| `setup_checklist` | Onboarding progress tracking |

## Security

- All tables have Row Level Security (RLS) enabled
- Owners can only access their own data
- Admins can access all data
- Stripe webhook validates signature on every request
- Service role key is only used server-side

## License

MIT
