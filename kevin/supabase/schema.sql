-- HOA Management SaaS - Supabase Schema
-- Run this in your Supabase SQL editor

-- Enable UUID extension
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- =====================
-- HOA Settings Table
-- =====================
CREATE TABLE IF NOT EXISTS hoa_settings (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  hoa_name TEXT NOT NULL DEFAULT 'My HOA',
  address TEXT,
  num_units INTEGER DEFAULT 0,
  default_monthly_dues NUMERIC(10,2) DEFAULT 0,
  default_due_day INTEGER DEFAULT 1,
  late_fee_amount NUMERIC(10,2) DEFAULT 0,
  late_fee_grace_days INTEGER DEFAULT 5,
  accounting_start_date DATE,
  payment_methods TEXT[] DEFAULT ARRAY['card'],
  setup_complete BOOLEAN DEFAULT FALSE,
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Insert default settings row
INSERT INTO hoa_settings (id) VALUES ('00000000-0000-0000-0000-000000000001')
ON CONFLICT (id) DO NOTHING;

-- =====================
-- Profiles Table
-- =====================
CREATE TABLE IF NOT EXISTS profiles (
  id UUID PRIMARY KEY REFERENCES auth.users(id) ON DELETE CASCADE,
  name TEXT NOT NULL DEFAULT '',
  email TEXT NOT NULL DEFAULT '',
  role TEXT NOT NULL DEFAULT 'owner' CHECK (role IN ('owner', 'admin')),
  unit_number TEXT,
  phone TEXT,
  move_in_date DATE,
  is_active BOOLEAN DEFAULT TRUE,
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- =====================
-- Dues Table
-- =====================
CREATE TABLE IF NOT EXISTS dues (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  owner_id UUID NOT NULL REFERENCES profiles(id) ON DELETE CASCADE,
  amount_due NUMERIC(10,2) NOT NULL DEFAULT 0,
  due_date DATE NOT NULL,
  status TEXT NOT NULL DEFAULT 'pending' CHECK (status IN ('pending', 'paid', 'partial', 'overdue', 'waived')),
  balance_remaining NUMERIC(10,2) NOT NULL DEFAULT 0,
  notes TEXT,
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- =====================
-- Payments Table
-- =====================
CREATE TABLE IF NOT EXISTS payments (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  owner_id UUID NOT NULL REFERENCES profiles(id) ON DELETE CASCADE,
  dues_id UUID REFERENCES dues(id) ON DELETE SET NULL,
  amount NUMERIC(10,2) NOT NULL,
  stripe_session_id TEXT,
  stripe_payment_intent_id TEXT,
  status TEXT NOT NULL DEFAULT 'pending' CHECK (status IN ('pending', 'completed', 'failed', 'refunded')),
  payment_method TEXT DEFAULT 'card',
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- =====================
-- Messages Table
-- =====================
CREATE TABLE IF NOT EXISTS messages (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  sender_id UUID NOT NULL REFERENCES profiles(id) ON DELETE CASCADE,
  recipient_id UUID REFERENCES profiles(id) ON DELETE CASCADE,
  subject TEXT,
  content TEXT NOT NULL,
  is_broadcast BOOLEAN DEFAULT FALSE,
  read_status BOOLEAN DEFAULT FALSE,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- =====================
-- Expenses Table
-- =====================
CREATE TABLE IF NOT EXISTS expenses (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  vendor_name TEXT NOT NULL,
  category TEXT NOT NULL,
  amount NUMERIC(10,2) NOT NULL,
  date DATE NOT NULL,
  notes TEXT,
  receipt_url TEXT,
  created_by UUID REFERENCES profiles(id) ON DELETE SET NULL,
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- =====================
-- Documents Table
-- =====================
CREATE TABLE IF NOT EXISTS documents (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  title TEXT NOT NULL,
  file_url TEXT NOT NULL,
  file_name TEXT,
  file_size INTEGER,
  category TEXT DEFAULT 'general',
  uploaded_by UUID REFERENCES profiles(id) ON DELETE SET NULL,
  is_public BOOLEAN DEFAULT TRUE,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- =====================
-- Setup Checklist Table
-- =====================
CREATE TABLE IF NOT EXISTS setup_checklist (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  step_key TEXT NOT NULL UNIQUE,
  step_label TEXT NOT NULL,
  completed BOOLEAN DEFAULT FALSE,
  completed_at TIMESTAMPTZ,
  sort_order INTEGER DEFAULT 0
);

INSERT INTO setup_checklist (step_key, step_label, sort_order) VALUES
  ('hoa_info', 'HOA Organization Settings', 1),
  ('admin_identity', 'Admin Identity Setup', 2),
  ('financial_setup', 'Financial Setup', 3),
  ('payment_config', 'Payment Configuration', 4),
  ('documents_upload', 'Initial Documents Upload', 5)
ON CONFLICT (step_key) DO NOTHING;

-- =====================
-- RLS Policies
-- =====================

-- Enable RLS on all tables
ALTER TABLE hoa_settings ENABLE ROW LEVEL SECURITY;
ALTER TABLE profiles ENABLE ROW LEVEL SECURITY;
ALTER TABLE dues ENABLE ROW LEVEL SECURITY;
ALTER TABLE payments ENABLE ROW LEVEL SECURITY;
ALTER TABLE messages ENABLE ROW LEVEL SECURITY;
ALTER TABLE expenses ENABLE ROW LEVEL SECURITY;
ALTER TABLE documents ENABLE ROW LEVEL SECURITY;
ALTER TABLE setup_checklist ENABLE ROW LEVEL SECURITY;

-- Helper function: is current user an admin?
CREATE OR REPLACE FUNCTION is_admin()
RETURNS BOOLEAN AS $$
  SELECT EXISTS (
    SELECT 1 FROM profiles
    WHERE id = auth.uid() AND role = 'admin'
  );
$$ LANGUAGE sql SECURITY DEFINER;

-- hoa_settings: admins read/write, owners read
CREATE POLICY "admins_all_hoa_settings" ON hoa_settings
  FOR ALL TO authenticated USING (is_admin()) WITH CHECK (is_admin());
CREATE POLICY "owners_read_hoa_settings" ON hoa_settings
  FOR SELECT TO authenticated USING (true);

-- profiles: users see own, admins see all
CREATE POLICY "users_read_own_profile" ON profiles
  FOR SELECT TO authenticated USING (id = auth.uid() OR is_admin());
CREATE POLICY "users_update_own_profile" ON profiles
  FOR UPDATE TO authenticated USING (id = auth.uid() OR is_admin());
CREATE POLICY "admins_insert_profile" ON profiles
  FOR INSERT TO authenticated WITH CHECK (is_admin() OR id = auth.uid());
CREATE POLICY "admins_delete_profile" ON profiles
  FOR DELETE TO authenticated USING (is_admin());

-- dues: owners see own, admins see all
CREATE POLICY "owners_read_own_dues" ON dues
  FOR SELECT TO authenticated USING (owner_id = auth.uid() OR is_admin());
CREATE POLICY "admins_insert_dues" ON dues
  FOR INSERT TO authenticated WITH CHECK (is_admin());
CREATE POLICY "admins_update_dues" ON dues
  FOR UPDATE TO authenticated USING (is_admin());
CREATE POLICY "admins_delete_dues" ON dues
  FOR DELETE TO authenticated USING (is_admin());

-- payments: owners see own, admins see all
CREATE POLICY "owners_read_own_payments" ON payments
  FOR SELECT TO authenticated USING (owner_id = auth.uid() OR is_admin());
CREATE POLICY "owners_insert_own_payment" ON payments
  FOR INSERT TO authenticated WITH CHECK (owner_id = auth.uid() OR is_admin());
CREATE POLICY "admins_update_payments" ON payments
  FOR UPDATE TO authenticated USING (is_admin());

-- messages: sender or recipient can read, admins see all
CREATE POLICY "users_read_messages" ON messages
  FOR SELECT TO authenticated USING (
    sender_id = auth.uid() OR
    recipient_id = auth.uid() OR
    is_broadcast = TRUE OR
    is_admin()
  );
CREATE POLICY "users_insert_messages" ON messages
  FOR INSERT TO authenticated WITH CHECK (
    sender_id = auth.uid() OR is_admin()
  );
CREATE POLICY "users_update_messages" ON messages
  FOR UPDATE TO authenticated USING (
    recipient_id = auth.uid() OR is_admin()
  );

-- expenses: admins only
CREATE POLICY "admins_all_expenses" ON expenses
  FOR ALL TO authenticated USING (is_admin()) WITH CHECK (is_admin());

-- documents: everyone reads public, admins write
CREATE POLICY "all_read_public_documents" ON documents
  FOR SELECT TO authenticated USING (is_public = TRUE OR is_admin());
CREATE POLICY "admins_insert_documents" ON documents
  FOR INSERT TO authenticated WITH CHECK (is_admin());
CREATE POLICY "admins_update_documents" ON documents
  FOR UPDATE TO authenticated USING (is_admin());
CREATE POLICY "admins_delete_documents" ON documents
  FOR DELETE TO authenticated USING (is_admin());

-- setup_checklist: admins only
CREATE POLICY "admins_all_checklist" ON setup_checklist
  FOR ALL TO authenticated USING (is_admin()) WITH CHECK (is_admin());
CREATE POLICY "all_read_checklist" ON setup_checklist
  FOR SELECT TO authenticated USING (true);

-- =====================
-- Triggers
-- =====================

-- Auto-create profile on signup
CREATE OR REPLACE FUNCTION handle_new_user()
RETURNS TRIGGER AS $$
BEGIN
  INSERT INTO profiles (id, email, name, role)
  VALUES (
    NEW.id,
    NEW.email,
    COALESCE(NEW.raw_user_meta_data->>'name', ''),
    COALESCE(NEW.raw_user_meta_data->>'role', 'owner')
  );
  RETURN NEW;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

CREATE OR REPLACE TRIGGER on_auth_user_created
  AFTER INSERT ON auth.users
  FOR EACH ROW EXECUTE FUNCTION handle_new_user();

-- Update updated_at timestamps
CREATE OR REPLACE FUNCTION update_updated_at()
RETURNS TRIGGER AS $$
BEGIN
  NEW.updated_at = NOW();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER update_profiles_updated_at BEFORE UPDATE ON profiles
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();
CREATE TRIGGER update_dues_updated_at BEFORE UPDATE ON dues
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();
CREATE TRIGGER update_payments_updated_at BEFORE UPDATE ON payments
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();
CREATE TRIGGER update_expenses_updated_at BEFORE UPDATE ON expenses
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();
CREATE TRIGGER update_hoa_settings_updated_at BEFORE UPDATE ON hoa_settings
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();

-- =====================
-- Storage
-- =====================
-- Run these in Supabase Dashboard > Storage:
-- 1. Create bucket: hoa-documents (public: false)
-- 2. Create bucket: receipts (public: false)
