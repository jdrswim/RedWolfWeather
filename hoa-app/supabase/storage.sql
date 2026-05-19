-- ============================================================
-- HOA Manager — Storage Setup
-- Run in Supabase SQL editor after schema.sql and rls.sql
-- OR create the bucket manually in the Supabase dashboard
-- ============================================================

-- Create the hoa-documents bucket
INSERT INTO storage.buckets (id, name, public)
VALUES ('hoa-documents', 'hoa-documents', true)
ON CONFLICT (id) DO NOTHING;

-- RLS policies for storage
CREATE POLICY "storage_admin_upload" ON storage.objects
  FOR INSERT WITH CHECK (
    bucket_id = 'hoa-documents'
    AND public.is_admin()
  );

CREATE POLICY "storage_authenticated_read" ON storage.objects
  FOR SELECT USING (
    bucket_id = 'hoa-documents'
    AND auth.uid() IS NOT NULL
  );

CREATE POLICY "storage_admin_delete" ON storage.objects
  FOR DELETE USING (
    bucket_id = 'hoa-documents'
    AND public.is_admin()
  );
