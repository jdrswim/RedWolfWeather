import { createServerSupabaseClient } from '@/lib/supabaseServer'
import { redirect } from 'next/navigation'
import { DollarSign, Plus } from 'lucide-react'
import { format } from 'date-fns'
import AssignDuesForm from './AssignDuesForm'

export default async function AdminDuesPage() {
  const supabase = await createServerSupabaseClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) redirect('/kevin/login')

  const { data: profile } = await supabase.from('profiles').select('role').eq('id', user.id).single()
  if (profile?.role !== 'admin') redirect('/kevin/owner')

  const [{ data: dues }, { data: owners }] = await Promise.all([
    supabase
      .from('dues')
      .select('*, profiles(name, unit_number)')
      .order('due_date', { ascending: false }),
    supabase.from('profiles').select('id, name, unit_number').eq('role', 'owner').order('name'),
  ])

  const statusColors: Record<string, string> = {
    pending: 'bg-yellow-100 text-yellow-700',
    paid: 'bg-green-100 text-green-700',
    partial: 'bg-blue-100 text-blue-700',
    overdue: 'bg-red-100 text-red-700',
    waived: 'bg-gray-100 text-gray-600',
  }

  return (
    <div className="max-w-6xl mx-auto space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-gray-900">Dues Management</h1>
          <p className="text-gray-500 text-sm mt-1">Assign and track HOA dues</p>
        </div>
      </div>

      <AssignDuesForm owners={owners || []} />

      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="px-6 py-4 border-b border-gray-100">
          <h2 className="font-semibold text-gray-900">All Dues</h2>
        </div>
        {!dues || dues.length === 0 ? (
          <div className="text-center py-16 text-gray-400">
            <DollarSign className="w-12 h-12 mx-auto mb-3 opacity-40" />
            <p className="font-medium">No dues assigned yet</p>
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="bg-gray-50 border-b border-gray-100">
                <tr>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Owner</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Amount</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Balance</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Due Date</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Status</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                {dues.map(d => (
                  <tr key={d.id} className="hover:bg-gray-50">
                    <td className="px-6 py-4">
                      <p className="font-medium text-gray-900">{(d.profiles as any)?.name || 'Unknown'}</p>
                      <p className="text-xs text-gray-400">Unit {(d.profiles as any)?.unit_number || '—'}</p>
                    </td>
                    <td className="px-6 py-4 font-semibold text-gray-900">${d.amount_due?.toFixed(2)}</td>
                    <td className="px-6 py-4">
                      <span className={d.balance_remaining > 0 ? 'text-red-600 font-semibold' : 'text-green-600 font-semibold'}>
                        ${d.balance_remaining?.toFixed(2)}
                      </span>
                    </td>
                    <td className="px-6 py-4 text-gray-600">{format(new Date(d.due_date), 'MMM d, yyyy')}</td>
                    <td className="px-6 py-4">
                      <span className={`px-2.5 py-0.5 rounded-full text-xs font-medium ${statusColors[d.status] || 'bg-gray-100 text-gray-600'}`}>
                        {d.status}
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
