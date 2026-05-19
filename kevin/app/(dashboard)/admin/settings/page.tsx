import { createServerSupabaseClient } from '@/lib/supabaseServer'
import { redirect } from 'next/navigation'
import SettingsForm from './SettingsForm'

export default async function AdminSettingsPage() {
  const supabase = await createServerSupabaseClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) redirect('/kevin/login')

  const { data: profile } = await supabase.from('profiles').select('role').eq('id', user.id).single()
  if (profile?.role !== 'admin') redirect('/kevin/owner')

  const [{ data: settings }, { data: checklist }] = await Promise.all([
    supabase.from('hoa_settings').select('*').eq('id', '00000000-0000-0000-0000-000000000001').single(),
    supabase.from('setup_checklist').select('*').order('sort_order'),
  ])

  return (
    <div className="max-w-3xl mx-auto space-y-6">
      <div>
        <h1 className="text-2xl font-bold text-gray-900">Settings & Onboarding</h1>
        <p className="text-gray-500 text-sm mt-1">Configure your HOA portal</p>
      </div>
      <SettingsForm settings={settings} checklist={checklist || []} />
    </div>
  )
}
