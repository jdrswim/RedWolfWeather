-- ============================================================
-- HOA Manager — Supabase Schema
-- Run this in your Supabase project SQL editor
-- ============================================================

-- Enable UUID extension
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- ============================================================
-- PROFILES (extends auth.users)
-- ============================================================
CREATE TABLE IF NOT EXISTS public.profiles (
  id          UUID PRIMARY KEY REFERENCES auth.users(id) ON DELETE CASCADE,
  name        TEXT,
  email       TEXT,
  role        TEXT NOT NULL DEFAULT 'owner' CHECK (role IN ('owner', 'admin')),
  unit_number TEXT,
  phone       TEXT,
  created_at  TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- Auto-create profile on signup
CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS TRIGGER AS $$
BEGIN
  INSERT INTO public.profiles (id, email, name, role, unit_number)
  VALUES (
    NEW.id,
    NEW.email,
    NEW.raw_user_meta_data ->> 'name',
    COALESCE(NEW.raw_user_meta_data ->> 'role', 'owner'),
    NEW.raw_user_meta_data ->> 'unit_number'
  )
  ON CONFLICT (id) DO NOTHING;
  RETURN NEW;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

DROP TRIGGER IF EXISTS on_auth_user_created ON auth.users;
CREATE TRIGGER on_auth_user_created
  AFTER INSERT ON auth.users
  FOR EACH ROW EXECUTE FUNCTION public.handle_new_user();

-- ============================================================
-- HOA SETTINGS
-- ============================================================
CREATE TABLE IF NOT EXISTS public.hoa_settings (
  id                    UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  hoa_name              TEXT NOT NULL DEFAULT '',
  address               TEXT DEFAULT '',
  city                  TEXT DEFAULT '',
  state                 TEXT DEFAULT '',
  zip                   TEXT DEFAULT '',
  num_units             INTEGER DEFAULT 0,
  default_monthly_dues  NUMERIC(10, 2),
  default_due_day       INTEGER DEFAULT 1 CHECK (default_due_day BETWEEN 1 AND 28),
  late_fee_amount       NUMERIC(10, 2),
  late_fee_days_grace   INTEGER DEFAULT 5,
  accounting_start_date DATE,
  payment_methods       TEXT[] DEFAULT ARRAY['card'],
  stripe_configured     BOOLEAN DEFAULT FALSE,
  onboarding_completed  BOOLEAN DEFAULT FALSE,
  created_at            TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at            TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- ============================================================
-- DUES
-- ============================================================
CREATE TABLE IF NOT EXISTS public.dues (
  id                UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  owner_id          UUID NOT NULL REFERENCES public.profiles(id) ON DELETE CASCADE,
  amount_due        NUMERIC(10, 2) NOT NULL,
  due_date          DATE NOT NULL,
  status            TEXT NOT NULL DEFAULT 'unpaid' CHECK (status IN ('unpaid', 'paid', 'overdue', 'partial')),
  balance_remaining NUMERIC(10, 2) NOT NULL,
  month_year        TEXT NOT NULL,
  notes             TEXT,
  created_at        TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_dues_owner_id ON public.dues(owner_id);
CREATE INDEX IF NOT EXISTS idx_dues_status ON public.dues(status);
CREATE INDEX IF NOT EXISTS idx_dues_due_date ON public.dues(due_date);

-- ============================================================
-- PAYMENTS
-- ============================================================
CREATE TABLE IF NOT EXISTS public.payments (
  id                         UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  owner_id                   UUID NOT NULL REFERENCES public.profiles(id) ON DELETE CASCADE,
  amount                     NUMERIC(10, 2) NOT NULL,
  stripe_session_id          TEXT,
  stripe_payment_intent_id   TEXT,
  status                     TEXT NOT NULL DEFAULT 'pending' CHECK (status IN ('pending', 'completed', 'failed', 'refunded')),
  payment_date               TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  due_id                     UUID REFERENCES public.dues(id) ON DELETE SET NULL,
  created_at                 TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_payments_owner_id ON public.payments(owner_id);
CREATE INDEX IF NOT EXISTS idx_payments_stripe_session ON public.payments(stripe_session_id);
CREATE INDEX IF NOT EXISTS idx_payments_status ON public.payments(status);

-- ============================================================
-- MESSAGES
-- ============================================================
CREATE TABLE IF NOT EXISTS public.messages (
  id           UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  sender_id    UUID NOT NULL REFERENCES public.profiles(id) ON DELETE CASCADE,
  recipient_id UUID REFERENCES public.profiles(id) ON DELETE CASCADE,
  content      TEXT NOT NULL,
  read_status  TEXT NOT NULL DEFAULT 'unread' CHECK (read_status IN ('read', 'unread')),
  is_broadcast BOOLEAN NOT NULL DEFAULT FALSE,
  created_at   TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_messages_sender ON public.messages(sender_id);
CREATE INDEX IF NOT EXISTS idx_messages_recipient ON public.messages(recipient_id);
CREATE INDEX IF NOT EXISTS idx_messages_broadcast ON public.messages(is_broadcast);

-- ============================================================
-- EXPENSES
-- ============================================================
CREATE TABLE IF NOT EXISTS public.expenses (
  id           UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  vendor_name  TEXT NOT NULL,
  category     TEXT NOT NULL,
  amount       NUMERIC(10, 2) NOT NULL,
  date         DATE NOT NULL,
  notes        TEXT,
  receipt_url  TEXT,
  created_by   UUID NOT NULL REFERENCES public.profiles(id) ON DELETE SET NULL,
  created_at   TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_expenses_date ON public.expenses(date);
CREATE INDEX IF NOT EXISTS idx_expenses_category ON public.expenses(category);

-- ============================================================
-- DOCUMENTS
-- ============================================================
CREATE TABLE IF NOT EXISTS public.documents (
  id          UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  title       TEXT NOT NULL,
  file_url    TEXT NOT NULL,
  file_name   TEXT NOT NULL,
  file_size   BIGINT,
  uploaded_by UUID REFERENCES public.profiles(id) ON DELETE SET NULL,
  category    TEXT,
  created_at  TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_documents_uploaded_by ON public.documents(uploaded_by);
