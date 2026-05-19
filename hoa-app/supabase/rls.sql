-- ============================================================
-- HOA Manager — Row Level Security Policies
-- Run AFTER schema.sql
-- ============================================================

-- Enable RLS on all tables
ALTER TABLE public.profiles ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.hoa_settings ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.dues ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.payments ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.messages ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.expenses ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.documents ENABLE ROW LEVEL SECURITY;

-- Helper: check if current user is admin
CREATE OR REPLACE FUNCTION public.is_admin()
RETURNS BOOLEAN AS $$
  SELECT EXISTS (
    SELECT 1 FROM public.profiles
    WHERE id = auth.uid() AND role = 'admin'
  );
$$ LANGUAGE sql SECURITY DEFINER STABLE;

-- ============================================================
-- PROFILES
-- ============================================================
-- Owners can read/update their own profile; admins can read all
CREATE POLICY "profiles_select_own" ON public.profiles
  FOR SELECT USING (auth.uid() = id OR public.is_admin());

CREATE POLICY "profiles_insert_own" ON public.profiles
  FOR INSERT WITH CHECK (auth.uid() = id OR public.is_admin());

CREATE POLICY "profiles_update_own" ON public.profiles
  FOR UPDATE USING (auth.uid() = id OR public.is_admin());

-- Admins can insert profiles for new owners
CREATE POLICY "profiles_admin_insert" ON public.profiles
  FOR INSERT WITH CHECK (public.is_admin());

-- ============================================================
-- HOA SETTINGS
-- ============================================================
-- Everyone can read settings (needed for HOA name display, etc.)
CREATE POLICY "hoa_settings_select_all" ON public.hoa_settings
  FOR SELECT USING (auth.uid() IS NOT NULL);

-- Only admins can insert/update
CREATE POLICY "hoa_settings_admin_write" ON public.hoa_settings
  FOR ALL USING (public.is_admin());

-- ============================================================
-- DUES
-- ============================================================
-- Owners see only their own dues; admins see all
CREATE POLICY "dues_owner_select" ON public.dues
  FOR SELECT USING (owner_id = auth.uid() OR public.is_admin());

-- Only admins can create/update/delete dues
CREATE POLICY "dues_admin_write" ON public.dues
  FOR INSERT WITH CHECK (public.is_admin());

CREATE POLICY "dues_admin_update" ON public.dues
  FOR UPDATE USING (public.is_admin());

CREATE POLICY "dues_admin_delete" ON public.dues
  FOR DELETE USING (public.is_admin());

-- ============================================================
-- PAYMENTS
-- ============================================================
-- Owners see only their own payments; admins see all
CREATE POLICY "payments_owner_select" ON public.payments
  FOR SELECT USING (owner_id = auth.uid() OR public.is_admin());

-- Owners can insert their own payments (Stripe checkout creates these)
CREATE POLICY "payments_owner_insert" ON public.payments
  FOR INSERT WITH CHECK (owner_id = auth.uid() OR public.is_admin());

-- Only admins (and service role via webhook) can update
CREATE POLICY "payments_admin_update" ON public.payments
  FOR UPDATE USING (public.is_admin());

-- ============================================================
-- MESSAGES
-- ============================================================
-- Users see messages they sent, received, or broadcasts
CREATE POLICY "messages_select" ON public.messages
  FOR SELECT USING (
    sender_id = auth.uid()
    OR recipient_id = auth.uid()
    OR is_broadcast = TRUE
    OR public.is_admin()
  );

-- Authenticated users can send messages
CREATE POLICY "messages_insert" ON public.messages
  FOR INSERT WITH CHECK (sender_id = auth.uid());

-- Users can mark their own received messages as read
CREATE POLICY "messages_update_read" ON public.messages
  FOR UPDATE USING (recipient_id = auth.uid() OR public.is_admin());

-- Only admins can delete messages
CREATE POLICY "messages_admin_delete" ON public.messages
  FOR DELETE USING (public.is_admin());

-- ============================================================
-- EXPENSES
-- ============================================================
-- Only admins can see and manage expenses
CREATE POLICY "expenses_admin_all" ON public.expenses
  FOR ALL USING (public.is_admin());

-- ============================================================
-- DOCUMENTS
-- ============================================================
-- All authenticated users can read documents
CREATE POLICY "documents_select_authenticated" ON public.documents
  FOR SELECT USING (auth.uid() IS NOT NULL);

-- Only admins can create/update/delete documents
CREATE POLICY "documents_admin_write" ON public.documents
  FOR INSERT WITH CHECK (public.is_admin());

CREATE POLICY "documents_admin_update" ON public.documents
  FOR UPDATE USING (public.is_admin());

CREATE POLICY "documents_admin_delete" ON public.documents
  FOR DELETE USING (public.is_admin());

-- ============================================================
-- REALTIME
-- ============================================================
-- Enable realtime for messages table
ALTER PUBLICATION supabase_realtime ADD TABLE public.messages;
