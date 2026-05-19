import { createServerSupabaseClient } from '@/lib/supabaseServer'
import { redirect } from 'next/navigation'
import { CreditCard } from 'lucide-react'
import { format } from 'date-fns'

export default async function AdminPaymentsPage() {
  const supabase = await createServerSupabaseClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) redirect('/kevin/login')

  const { data: profile } = await supabase.from('profiles').select('role').eq('id', user.id).single()
  if (profile?.role !== 'admin') redirect('/kevin/owner')

  const { data: payments } = await supabase
    .from('payments')
    .select('*, profiles(name, unit_number)')
    .order('created_at', { ascending: false })

  const totalCollected = payments?.filter(p => p.status === 'completed').reduce((s, p) => s + p.amount, 0) || 0
  const pendingAmount = payments?.filter(p => p.status === 'pending').reduce((s, p) => s + p.amount, 0) || 0

  return (
    <div className="max-w-6xl mx-auto space-y-6">
      <div>
        <h1 className="text-2xl font-bold text-gray-900">Payments</h1>
        <p className="text-gray-500 text-sm mt-1">All HOA payment records</p>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-5">
        <div className="bg-green-50 border border-green-100 rounded-2xl p-6">
          <p className="text-sm text-gray-500 mb-1">Total Collected</p>
          <p className="text-3xl font-bold text-green-700">${totalCollected.toFixed(2)}</p>
        </div>
        <div className="bg-yellow-50 border border-yellow-100 rounded-2xl p-6">
          <p className="text-sm text-gray-500 mb-1">Pending</p>
          <p className="text-3xl font-bold text-yellow-600">${pendingAmount.toFixed(2)}</p>
        </div>
      </div>

      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="px-6 py-4 border-b border-gray-100">
          <h2 className="font-semibold text-gray-900">All Transactions</h2>
        </div>
        {!payments || payments.length === 0 ? (
          <div className="text-center py-16 text-gray-400">
            <CreditCard className="w-12 h-12 mx-auto mb-3 opacity-40" />
            <p className="font-medium">No payments yet</p>
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="bg-gray-50 border-b border-gray-100">
                <tr>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Date</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Owner</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Amount</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Status</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Session ID</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                {payments.map(p => (
                  <tr key={p.id} className="hover:bg-gray-50">
                    <td className="px-6 py-4 text-gray-600">{format(new Date(p.created_at), 'MMM d, yyyy')}</td>
                    <td className="px-6 py-4">
                      <p className="font-medium text-gray-900">{(p.profiles as any)?.name || '—'}</p>
                      <p className="text-xs text-gray-400">Unit {(p.profiles as any)?.unit_number || '—'}</p>
                    </td>
                    <td className="px-6 py-4 font-semibold text-gray-900">${p.amount?.toFixed(2)}</td>
                    <td className="px-6 py-4">
                      <span className={`px-2.5 py-0.5 rounded-full text-xs font-medium ${
                        p.status === 'completed' ? 'bg-green-100 text-green-700' :
                        p.status === 'failed' ? 'bg-red-100 text-red-700' :
                        'bg-yellow-100 text-yellow-700'
                      }`}>
                        {p.status}
                      </span>
                    </td>
                    <td className="px-6 py-4 text-gray-400 text-xs font-mono">
                      {p.stripe_session_id ? p.stripe_session_id.slice(0, 24) + '...' : '—'}
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
