import { createServerSupabaseClient } from '@/lib/supabaseServer'
import { redirect } from 'next/navigation'
import StatCard from '@/components/StatCard'
import { DollarSign, CreditCard, MessageSquare, FileText } from 'lucide-react'
import Link from 'next/link'
import { format } from 'date-fns'

export default async function OwnerDashboard() {
  const supabase = await createServerSupabaseClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) redirect('/kevin/login')

  const { data: profile } = await supabase
    .from('profiles')
    .select('*')
    .eq('id', user.id)
    .single()

  const [
    { data: dues },
    { data: payments },
    { data: messages },
    { data: documents },
    { data: settings },
  ] = await Promise.all([
    supabase.from('dues').select('*').eq('owner_id', user.id).order('due_date', { ascending: false }),
    supabase.from('payments').select('*').eq('owner_id', user.id).order('created_at', { ascending: false }),
    supabase.from('messages').select('*').or(`recipient_id.eq.${user.id},is_broadcast.eq.true`).eq('read_status', false),
    supabase.from('documents').select('id').eq('is_public', true),
    supabase.from('hoa_settings').select('*').eq('id', '00000000-0000-0000-0000-000000000001').single(),
  ])

  const totalBalance = dues?.reduce((sum, d) => sum + (d.balance_remaining || 0), 0) || 0
  const pendingDue = dues?.find(d => d.status === 'pending' || d.status === 'overdue')
  const unreadMessages = messages?.length || 0

  return (
    <div className="max-w-5xl mx-auto space-y-8">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-gray-900">
          Welcome back, {profile?.name || 'Homeowner'}
        </h1>
        <p className="text-gray-500 text-sm mt-1">
          {settings?.data?.hoa_name || "Kevin's HOA"} • Unit {profile?.unit_number || '—'}
        </p>
      </div>

      {/* Stats */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5">
        <StatCard label="Balance Due" value={`$${totalBalance.toFixed(2)}`} icon={DollarSign} color={totalBalance > 0 ? 'red' : 'green'} />
        <StatCard label="Payments Made" value={payments?.filter(p => p.status === 'completed').length || 0} icon={CreditCard} color="blue" />
        <StatCard label="Unread Messages" value={unreadMessages} icon={MessageSquare} color={unreadMessages > 0 ? 'yellow' : 'blue'} />
        <StatCard label="Documents" value={documents?.length || 0} icon={FileText} color="purple" />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Current Dues */}
        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6">
          <div className="flex items-center justify-between mb-5">
            <h2 className="text-lg font-semibold text-gray-900">Current Dues</h2>
            <Link href="/kevin/owner/dues" className="text-sm text-blue-600 hover:text-blue-700 font-medium">View all</Link>
          </div>
          {!pendingDue ? (
            <div className="text-center py-6 text-gray-400">
              <DollarSign className="w-10 h-10 mx-auto mb-2 opacity-40" />
              <p className="text-sm">All dues paid — great job!</p>
            </div>
          ) : (
            <div className="space-y-4">
              <div className={`p-4 rounded-xl border-2 ${pendingDue.status === 'overdue' ? 'border-red-200 bg-red-50' : 'border-yellow-200 bg-yellow-50'}`}>
                <div className="flex items-center justify-between">
                  <div>
                    <p className="text-sm font-medium text-gray-700">
                      Due {format(new Date(pendingDue.due_date), 'MMMM d, yyyy')}
                    </p>
                    <p className="text-xs text-gray-500 mt-0.5 capitalize">{pendingDue.status}</p>
                  </div>
                  <div className="text-right">
                    <p className="text-xl font-bold text-gray-900">${pendingDue.balance_remaining?.toFixed(2)}</p>
                    <p className="text-xs text-gray-400">of ${pendingDue.amount_due?.toFixed(2)}</p>
                  </div>
                </div>
                <Link
                  href={`/kevin/owner/dues`}
                  className="mt-3 block w-full text-center py-2.5 px-4 bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold rounded-xl transition-colors"
                >
                  Pay Now
                </Link>
              </div>
            </div>
          )}
        </div>

        {/* Recent Payments */}
        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6">
          <div className="flex items-center justify-between mb-5">
            <h2 className="text-lg font-semibold text-gray-900">Recent Payments</h2>
            <Link href="/kevin/owner/payments" className="text-sm text-blue-600 hover:text-blue-700 font-medium">View all</Link>
          </div>
          {!payments || payments.length === 0 ? (
            <div className="text-center py-6 text-gray-400">
              <CreditCard className="w-10 h-10 mx-auto mb-2 opacity-40" />
              <p className="text-sm">No payments yet</p>
            </div>
          ) : (
            <div className="space-y-3">
              {payments.slice(0, 4).map(p => (
                <div key={p.id} className="flex items-center justify-between py-2.5 border-b border-gray-50 last:border-0">
                  <div>
                    <p className="text-sm font-medium text-gray-700">${p.amount?.toFixed(2)}</p>
                    <p className="text-xs text-gray-400">{format(new Date(p.created_at), 'MMM d, yyyy')}</p>
                  </div>
                  <span className={`px-2.5 py-0.5 rounded-full text-xs font-medium ${
                    p.status === 'completed' ? 'bg-green-100 text-green-700' :
                    p.status === 'failed' ? 'bg-red-100 text-red-700' :
                    'bg-yellow-100 text-yellow-700'
                  }`}>
                    {p.status}
                  </span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
