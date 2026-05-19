import { createServerSupabaseClient } from '@/lib/supabaseServer'
import { redirect } from 'next/navigation'
import { DollarSign, CheckCircle, AlertCircle, Clock } from 'lucide-react'
import { format } from 'date-fns'
import PayDuesButton from './PayDuesButton'

export default async function OwnerDuesPage() {
  const supabase = await createServerSupabaseClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) redirect('/kevin/login')

  const { data: dues } = await supabase
    .from('dues')
    .select('*')
    .eq('owner_id', user.id)
    .order('due_date', { ascending: false })

  const statusConfig: Record<string, { icon: typeof CheckCircle; color: string; badge: string }> = {
    paid: { icon: CheckCircle, color: 'text-green-500', badge: 'bg-green-100 text-green-700' },
    pending: { icon: Clock, color: 'text-yellow-500', badge: 'bg-yellow-100 text-yellow-700' },
    partial: { icon: Clock, color: 'text-blue-500', badge: 'bg-blue-100 text-blue-700' },
    overdue: { icon: AlertCircle, color: 'text-red-500', badge: 'bg-red-100 text-red-700' },
    waived: { icon: CheckCircle, color: 'text-gray-400', badge: 'bg-gray-100 text-gray-600' },
  }

  const totalBalance = dues?.reduce((sum, d) => sum + (d.balance_remaining || 0), 0) || 0

  return (
    <div className="max-w-4xl mx-auto space-y-6">
      <div>
        <h1 className="text-2xl font-bold text-gray-900">My Dues</h1>
        <p className="text-gray-500 text-sm mt-1">View and pay your HOA dues</p>
      </div>

      {/* Balance summary */}
      <div className={`rounded-2xl p-6 ${totalBalance > 0 ? 'bg-red-50 border border-red-100' : 'bg-green-50 border border-green-100'}`}>
        <div className="flex items-center justify-between">
          <div>
            <p className="text-sm font-medium text-gray-500 mb-1">Total Balance Due</p>
            <p className={`text-4xl font-bold ${totalBalance > 0 ? 'text-red-600' : 'text-green-600'}`}>
              ${totalBalance.toFixed(2)}
            </p>
          </div>
          <DollarSign className={`w-12 h-12 ${totalBalance > 0 ? 'text-red-300' : 'text-green-300'}`} />
        </div>
        {totalBalance === 0 && (
          <p className="text-green-600 text-sm font-medium mt-2">All dues are paid — great job!</p>
        )}
      </div>

      {/* Dues list */}
      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="px-6 py-4 border-b border-gray-100">
          <h2 className="font-semibold text-gray-900">Dues History</h2>
        </div>
        {!dues || dues.length === 0 ? (
          <div className="text-center py-16 text-gray-400">
            <DollarSign className="w-12 h-12 mx-auto mb-3 opacity-40" />
            <p className="font-medium">No dues assigned yet</p>
          </div>
        ) : (
          <ul className="divide-y divide-gray-50">
            {dues.map(d => {
              const config = statusConfig[d.status] || statusConfig.pending
              const Icon = config.icon
              const canPay = d.status !== 'paid' && d.status !== 'waived' && d.balance_remaining > 0
              return (
                <li key={d.id} className="flex items-center gap-4 px-6 py-5">
                  <Icon className={`w-6 h-6 flex-shrink-0 ${config.color}`} />
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center gap-2 mb-0.5">
                      <p className="font-medium text-gray-900">
                        Due {format(new Date(d.due_date), 'MMMM d, yyyy')}
                      </p>
                      <span className={`px-2 py-0.5 rounded-full text-xs font-medium ${config.badge}`}>
                        {d.status}
                      </span>
                    </div>
                    <p className="text-sm text-gray-500">
                      ${d.amount_due?.toFixed(2)} total
                      {d.balance_remaining < d.amount_due && d.balance_remaining > 0 && (
                        <> • ${d.balance_remaining?.toFixed(2)} remaining</>
                      )}
                    </p>
                    {d.notes && <p className="text-xs text-gray-400 mt-0.5">{d.notes}</p>}
                  </div>
                  <div className="text-right flex-shrink-0">
                    {canPay ? (
                      <PayDuesButton duesId={d.id} amount={d.balance_remaining} />
                    ) : (
                      <span className="text-sm font-semibold text-green-600">
                        {d.status === 'paid' ? 'Paid' : d.status === 'waived' ? 'Waived' : ''}
                      </span>
                    )}
                  </div>
                </li>
              )
            })}
          </ul>
        )}
      </div>
    </div>
  )
}
