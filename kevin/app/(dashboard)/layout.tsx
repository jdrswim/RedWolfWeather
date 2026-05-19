import { redirect } from 'next/navigation'
import { createServerSupabaseClient } from '@/lib/supabaseServer'
import Sidebar from '@/components/Sidebar'

export default async function DashboardLayout({ children }: { children: React.ReactNode }) {
  const supabase = await createServerSupabaseClient()
  const { data: { user } } = await supabase.auth.getUser()

  if (!user) redirect('/kevin/login')

  const { data: profile } = await supabase
    .from('profiles')
    .select('name, role, unit_number')
    .eq('id', user.id)
    .single()

  return (
    <div className="flex min-h-screen bg-gray-50">
      <Sidebar
        role={(profile?.role as 'admin' | 'owner') || 'owner'}
        userName={profile?.name || user.email || 'User'}
        unitNumber={profile?.unit_number}
      />
      <main className="flex-1 min-w-0 overflow-auto">
        <div className="p-6 lg:p-8 pt-16 lg:pt-8">
          {children}
        </div>
      </main>
    </div>
  )
}
