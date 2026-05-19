import { redirect } from 'next/navigation'
import { createServerSupabaseClient } from '@/lib/supabaseServer'

export default async function HomePage() {
  const supabase = await createServerSupabaseClient()
  const { data: { user } } = await supabase.auth.getUser()

  if (user) {
    const { data: profile } = await supabase
      .from('profiles')
      .select('role')
      .eq('id', user.id)
      .single()
    redirect(profile?.role === 'admin' ? '/kevin/admin' : '/kevin/owner')
  }

  redirect('/kevin/login')
}
