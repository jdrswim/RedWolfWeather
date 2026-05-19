import { createServerSupabaseClient } from '@/lib/supabaseServer'
import { redirect } from 'next/navigation'
import MessagingPanel from '@/components/MessagingPanel'

export default async function AdminMessagesPage() {
  const supabase = await createServerSupabaseClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) redirect('/kevin/login')

  const { data: profile } = await supabase.from('profiles').select('*').eq('id', user.id).single()
  if (profile?.role !== 'admin') redirect('/kevin/owner')

  const { data: owners } = await supabase
    .from('profiles')
    .select('id, name, unit_number')
    .eq('role', 'owner')
    .order('name')

  const { data: messages } = await supabase
    .from('messages')
    .select('*, sender:profiles!sender_id(name), recipient:profiles!recipient_id(name)')
    .order('created_at', { ascending: false })
    .limit(50)

  return (
    <div className="max-w-5xl mx-auto space-y-6">
      <div>
        <h1 className="text-2xl font-bold text-gray-900">Messages</h1>
        <p className="text-gray-500 text-sm mt-1">Send direct messages or broadcast to all owners</p>
      </div>
      <MessagingPanel
        currentUser={{ id: user.id, name: profile?.name || '', role: 'admin' }}
        recipients={owners || []}
        messages={messages || []}
      />
    </div>
  )
}
