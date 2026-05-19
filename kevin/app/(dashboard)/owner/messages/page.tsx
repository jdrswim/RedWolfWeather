import { createServerSupabaseClient } from '@/lib/supabaseServer'
import { redirect } from 'next/navigation'
import MessagingPanel from '@/components/MessagingPanel'

export default async function OwnerMessagesPage() {
  const supabase = await createServerSupabaseClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) redirect('/kevin/login')

  const { data: profile } = await supabase.from('profiles').select('*').eq('id', user.id).single()

  const { data: admins } = await supabase
    .from('profiles')
    .select('id, name')
    .eq('role', 'admin')

  const { data: messages } = await supabase
    .from('messages')
    .select('*, sender:profiles!sender_id(name)')
    .or(`sender_id.eq.${user.id},recipient_id.eq.${user.id},is_broadcast.eq.true`)
    .order('created_at', { ascending: false })
    .limit(50)

  return (
    <div className="max-w-5xl mx-auto space-y-6">
      <div>
        <h1 className="text-2xl font-bold text-gray-900">Messages</h1>
        <p className="text-gray-500 text-sm mt-1">Contact your HOA administration</p>
      </div>
      <MessagingPanel
        currentUser={{ id: user.id, name: profile?.name || '', role: 'owner' }}
        recipients={admins || []}
        messages={messages || []}
      />
    </div>
  )
}
