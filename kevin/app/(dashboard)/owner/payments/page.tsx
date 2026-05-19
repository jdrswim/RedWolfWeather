import { createServerSupabaseClient } from '@/lib/supabaseServer'
import { redirect } from 'next/navigation'
import { CreditCard, CheckCircle, AlertCircle } from 'lucide-react'
import { format } from 'date-fns'

export default async function OwnerPaymentsPage({
  searchParams,
}: {
  searchParams: { success?: string; cancelled?: string }
}) {
  const supabase = await createServerSupabaseClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) redirect('/kevin/login')

  const { data: payments } = await supabase
    .from('payments')
    .select('*')
    .eq('owner_id', user.id)
    .order('created_at', { ascending: false })

  const totalPaid = payments?.filter(p => p.status === 'completed').reduce((s, p) => s + p.amount, 0) || 0

  return (
    <div className="max-w-4xl mx-auto space-y-6">
      <div>
        <h1 className="text-2xl font-bold text-gray-900">Payment History</h1>
        <p className="text-gray-500 text-sm mt-1">All your HOA payment records</p>
      </div>

      {searchParams.success === 'true' && (
        <div className="flex items-center gap-3 p-4 bg-green-50 border border-green-200 rounded-2xl text-green-700">
          <CheckCircle className="w-5 h-5 flex-shrink-0" />
          <p className="font-medium">Payment successful! Your dues have been updated.</p>
        </div>
      )}
      {searchParams.cancelled === 'true' && (
        <div className="flex items-center gap-3 p-4 bg-yellow-50 border border-yellow-200 rounded-2xl text-yellow-700">
          <AlertCircle className="w-5 h-5 flex-shrink-0" />
          <p className="font-medium">Payment cancelled. Your dues have not been updated.</p>
        </div>
      )}

      <div className="bg-blue-50 border border-blue-100 rounded-2xl p-6">
        <p className="text-sm text-gray-500 mb-1">Total Paid (All Time)</p>
        <p className="text-3xl font-bold text-blue-700">${totalPaid.toFixed(2)}</p>
      </div>

      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="px-6 py-4 border-b border-gray-100">
          <h2 className="font-semibold text-gray-900">Transactions</h2>
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
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Amount</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Method</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Status</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Reference</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                {payments.map(p => (
                  <tr key={p.id} className="hover:bg-gray-50">
                    <td className="px-6 py-4 text-gray-600">{format(new Date(p.created_at), 'MMM d, yyyy')}</td>
                    <td className="px-6 py-4 font-semibold text-gray-900">${p.amount?.toFixed(2)}</td>
                    <td className="px-6 py-4 text-gray-500 capitalize">{p.payment_method || 'card'}</td>
                    <td className="px-6 py-4">
                      <span className={`px-2.5 py-0.5 rounded-full text-xs font-medium ${
                        p.status === 'completed' ? 'bg-green-100 text-green-700' :
                        p.status === 'failed' ? 'bg-red-100 text-red-700' :
                        'bg-yellow-100 text-yellow-700'
                      }`}>
                        {p.status}
                      </span>
                    </td>
                    <td className="px-6 py-4 text-gray-400 text-xs font-mono truncate max-w-[140px]">
                      {p.stripe_session_id ? p.stripe_session_id.slice(0, 20) + '...' : '—'}
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
