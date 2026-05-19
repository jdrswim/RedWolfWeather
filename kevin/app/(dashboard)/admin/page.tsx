import { createServerSupabaseClient } from '@/lib/supabaseServer'
import { redirect } from 'next/navigation'
import StatCard from '@/components/StatCard'
import SetupChecklist from '@/components/SetupChecklist'
import { DollarSign, Users, CreditCard, TrendingUp, AlertCircle } from 'lucide-react'
import { format } from 'date-fns'

export default async function AdminDashboard() {
  const supabase = await createServerSupabaseClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) redirect('/kevin/login')

  const { data: profile } = await supabase.from('profiles').select('role').eq('id', user.id).single()
  if (profile?.role !== 'admin') redirect('/kevin/owner')

  const [
    { data: owners },
    { data: dues },
    { data: payments },
    { data: expenses },
    { data: checklist },
    { data: settings },
  ] = await Promise.all([
    supabase.from('profiles').select('id').eq('role', 'owner'),
    supabase.from('dues').select('*'),
    supabase.from('payments').select('amount, status, created_at').order('created_at', { ascending: false }),
    supabase.from('expenses').select('amount, date').order('date', { ascending: false }),
    supabase.from('setup_checklist').select('*').order('sort_order'),
    supabase.from('hoa_settings').select('*').eq('id', '00000000-0000-0000-0000-000000000001').single(),
  ])

  const totalDue = dues?.reduce((sum, d) => sum + (d.balance_remaining || 0), 0) || 0
  const totalCollected = payments?.filter(p => p.status === 'completed').reduce((sum, p) => sum + p.amount, 0) || 0
  const overdueCount = dues?.filter(d => d.status === 'overdue').length || 0
  const thisMonthExpenses = expenses?.filter(e => {
    const d = new Date(e.date)
    const now = new Date()
    return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear()
  }).reduce((sum, e) => sum + e.amount, 0) || 0

  const recentPayments = payments?.slice(0, 5) || []
  const overdueDues = dues?.filter(d => d.status === 'overdue') || []

  return (
    <div className="max-w-7xl mx-auto space-y-8">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-gray-900">Admin Dashboard</h1>
          <p className="text-gray-500 text-sm mt-1">{settings?.data?.hoa_name || "Kevin's HOA"} • {format(new Date(), 'MMMM yyyy')}</p>
        </div>
      </div>

      {/* Stats */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5">
        <StatCard label="Total Owners" value={owners?.length || 0} icon={Users} color="blue" />
        <StatCard label="Balance Due" value={`$${totalDue.toFixed(2)}`} icon={DollarSign} color="yellow" />
        <StatCard label="Collected (All Time)" value={`$${totalCollected.toFixed(2)}`} icon={CreditCard} color="green" />
        <StatCard label="This Month Expenses" value={`$${thisMonthExpenses.toFixed(2)}`} icon={TrendingUp} color="purple" />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Setup Checklist */}
        {checklist && checklist.some(c => !c.completed) && (
          <div className="lg:col-span-1">
            <SetupChecklist items={checklist} />
          </div>
        )}

        {/* Overdue Dues */}
        <div className={`bg-white rounded-2xl shadow-sm border border-gray-100 p-6 ${checklist && checklist.some(c => !c.completed) ? 'lg:col-span-2' : 'lg:col-span-3'}`}>
          <div className="flex items-center gap-2 mb-5">
            <AlertCircle className="w-5 h-5 text-red-500" />
            <h2 className="text-lg font-semibold text-gray-900">Overdue Dues ({overdueCount})</h2>
          </div>
          {overdueDues.length === 0 ? (
            <div className="text-center py-8 text-gray-400">
              <CreditCard className="w-10 h-10 mx-auto mb-2 opacity-40" />
              <p className="text-sm">No overdue dues</p>
            </div>
          ) : (
            <div className="space-y-3">
              {overdueDues.slice(0, 6).map(d => (
                <div key={d.id} className="flex items-center justify-between p-3 bg-red-50 rounded-xl">
                  <div>
                    <p className="text-sm font-medium text-gray-700">Due {format(new Date(d.due_date), 'MMM d, yyyy')}</p>
                  </div>
                  <div className="text-right">
                    <p className="text-sm font-bold text-red-600">${d.balance_remaining?.toFixed(2)}</p>
                    <p className="text-xs text-gray-400">of ${d.amount_due?.toFixed(2)}</p>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      {/* Recent Payments */}
      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6">
        <h2 className="text-lg font-semibold text-gray-900 mb-5">Recent Payments</h2>
        {recentPayments.length === 0 ? (
          <div className="text-center py-8 text-gray-400">
            <CreditCard className="w-10 h-10 mx-auto mb-2 opacity-40" />
            <p className="text-sm">No payments yet</p>
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-gray-500 border-b border-gray-100">
                  <th className="pb-3 font-medium">Date</th>
                  <th className="pb-3 font-medium">Amount</th>
                  <th className="pb-3 font-medium">Status</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                {recentPayments.map(p => (
                  <tr key={p.created_at} className="hover:bg-gray-50">
                    <td className="py-3 text-gray-600">{format(new Date(p.created_at), 'MMM d, yyyy')}</td>
                    <td className="py-3 font-semibold text-gray-900">${p.amount?.toFixed(2)}</td>
                    <td className="py-3">
                      <span className={`px-2.5 py-0.5 rounded-full text-xs font-medium ${
                        p.status === 'completed' ? 'bg-green-100 text-green-700' :
                        p.status === 'failed' ? 'bg-red-100 text-red-700' :
                        'bg-yellow-100 text-yellow-700'
                      }`}>
                        {p.status}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  )
}
