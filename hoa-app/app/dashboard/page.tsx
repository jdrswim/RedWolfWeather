import { redirect } from 'next/navigation'
import { createClient } from '@/lib/supabaseServer'

export default async function DashboardRootPage() {
  const supabase = await createClient()
  const { data: { user } } = await supabase.auth.getUser()

  if (!user) {
    redirect('/login')
  }

  const { data: profile } = await supabase
    .from('profiles')
    .select('role')
    .eq('id', user.id)
    .single()

  if (profile?.role === 'admin') {
    redirect('/dashboard/admin')
  } else {
    redirect('/dashboard/owner')
  }
}
