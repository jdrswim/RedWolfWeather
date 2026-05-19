import { createServerSupabaseClient } from '@/lib/supabaseServer'
import { redirect } from 'next/navigation'
import { format } from 'date-fns'
import AddExpenseForm from './AddExpenseForm'
import { ReceiptIcon } from 'lucide-react'

const CATEGORIES = ['Insurance', 'Utilities', 'Landscaping', 'Repairs', 'Management', 'Legal', 'Accounting', 'Other']

export default async function AdminExpensesPage() {
  const supabase = await createServerSupabaseClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) redirect('/kevin/login')

  const { data: profile } = await supabase.from('profiles').select('role').eq('id', user.id).single()
  if (profile?.role !== 'admin') redirect('/kevin/owner')

  const { data: expenses } = await supabase
    .from('expenses')
    .select('*')
    .order('date', { ascending: false })

  const totalExpenses = expenses?.reduce((s, e) => s + e.amount, 0) || 0
  const thisMonth = expenses?.filter(e => {
    const d = new Date(e.date)
    const n = new Date()
    return d.getMonth() === n.getMonth() && d.getFullYear() === n.getFullYear()
  }).reduce((s, e) => s + e.amount, 0) || 0

  return (
    <div className="max-w-6xl mx-auto space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-gray-900">Expenses</h1>
          <p className="text-gray-500 text-sm mt-1">Track HOA operating expenses</p>
        </div>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-5">
        <div className="bg-white rounded-2xl border border-gray-100 shadow-sm p-6">
          <p className="text-sm text-gray-500 mb-1">Total Expenses</p>
          <p className="text-3xl font-bold text-gray-900">${totalExpenses.toFixed(2)}</p>
        </div>
        <div className="bg-white rounded-2xl border border-gray-100 shadow-sm p-6">
          <p className="text-sm text-gray-500 mb-1">This Month</p>
          <p className="text-3xl font-bold text-gray-900">${thisMonth.toFixed(2)}</p>
        </div>
      </div>

      <AddExpenseForm />

      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="px-6 py-4 border-b border-gray-100">
          <h2 className="font-semibold text-gray-900">Expense Log</h2>
        </div>
        {!expenses || expenses.length === 0 ? (
          <div className="text-center py-16 text-gray-400">
            <ReceiptIcon className="w-12 h-12 mx-auto mb-3 opacity-40" />
            <p className="font-medium">No expenses recorded</p>
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="bg-gray-50 border-b border-gray-100">
                <tr>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Date</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Vendor</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Category</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Amount</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Notes</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                {expenses.map(e => (
                  <tr key={e.id} className="hover:bg-gray-50">
                    <td className="px-6 py-4 text-gray-600">{format(new Date(e.date), 'MMM d, yyyy')}</td>
                    <td className="px-6 py-4 font-medium text-gray-900">{e.vendor_name}</td>
                    <td className="px-6 py-4">
                      <span className="px-2.5 py-0.5 rounded-full bg-gray-100 text-gray-600 text-xs font-medium">{e.category}</span>
                    </td>
                    <td className="px-6 py-4 font-semibold text-gray-900">${e.amount?.toFixed(2)}</td>
                    <td className="px-6 py-4 text-gray-500 max-w-xs truncate">{e.notes || '—'}</td>
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
