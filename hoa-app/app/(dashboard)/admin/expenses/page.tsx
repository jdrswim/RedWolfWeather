'use client'

import { useEffect, useState, useCallback } from 'react'
import { createClient } from '@/lib/supabaseClient'
import AdminSidebar from '@/components/layout/AdminSidebar'
import PageHeader from '@/components/layout/PageHeader'
import Button from '@/components/ui/Button'
import Modal from '@/components/ui/Modal'
import { formatCurrency, formatDate } from '@/lib/utils'
import { Plus, Receipt, Search } from 'lucide-react'
import type { Expense } from '@/types'

const CATEGORIES = [
  'Insurance', 'Utilities', 'Landscaping', 'Maintenance', 'Repairs',
  'Management', 'Legal', 'Accounting', 'Reserve Fund', 'Other',
]

export default function ExpensesPage() {
  const [expenses, setExpenses] = useState<Expense[]>([])
  const [filtered, setFiltered] = useState<Expense[]>([])
  const [search, setSearch] = useState('')
  const [categoryFilter, setCategoryFilter] = useState('all')
  const [loading, setLoading] = useState(true)
  const [showAdd, setShowAdd] = useState(false)
  const [hoaName, setHoaName] = useState('')
  const supabase = createClient()

  const fetchExpenses = useCallback(async () => {
    const { data } = await supabase.from('expenses').select('*').order('date', { ascending: false })
    setExpenses(data ?? [])
    setLoading(false)
  }, [supabase])

  useEffect(() => {
    fetchExpenses()
    supabase.from('hoa_settings').select('hoa_name').single().then(({ data }) => { if (data) setHoaName(data.hoa_name) })
  }, [fetchExpenses, supabase])

  useEffect(() => {
    let r = expenses
    if (categoryFilter !== 'all') r = r.filter((e) => e.category === categoryFilter)
    if (search) {
      const q = search.toLowerCase()
      r = r.filter((e) => e.vendor_name.toLowerCase().includes(q) || e.category.toLowerCase().includes(q) || e.notes?.toLowerCase().includes(q))
    }
    setFiltered(r)
  }, [expenses, search, categoryFilter])

  const totalShown = filtered.reduce((s, e) => s + e.amount, 0)
  const totalAll = expenses.reduce((s, e) => s + e.amount, 0)

  const byCategory = CATEGORIES.map((c) => ({
    category: c,
    total: expenses.filter((e) => e.category === c).reduce((s, e) => s + e.amount, 0),
  })).filter((c) => c.total > 0)

  return (
    <div className="flex min-h-screen bg-gray-50">
      <AdminSidebar hoaName={hoaName} />
      <main className="flex-1 ml-64 p-8">
        <div className="max-w-6xl mx-auto">
          <PageHeader
            title="Expenses"
            description="Track HOA operating expenses and reserves"
            action={
              <Button onClick={() => setShowAdd(true)}>
                <Plus className="h-4 w-4" />
                Record expense
              </Button>
            }
          />

          <div className="grid grid-cols-1 xl:grid-cols-4 gap-6">
            {/* Main content */}
            <div className="xl:col-span-3">
              <div className="flex gap-3 mb-5">
                <div className="relative flex-1">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-gray-400" />
                  <input
                    type="text"
                    placeholder="Search vendor or category…"
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                    className="w-full pl-9 pr-4 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"
                  />
                </div>
                <select
                  value={categoryFilter}
                  onChange={(e) => setCategoryFilter(e.target.value)}
                  className="px-3 py-2.5 border border-gray-200 rounded-lg text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="all">All categories</option>
                  {CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>

              <div className="bg-white rounded-xl border border-gray-100 shadow-sm overflow-hidden">
                <div className="px-6 py-3 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
                  <span className="text-xs text-gray-500 font-medium">{filtered.length} expense{filtered.length !== 1 ? 's' : ''}</span>
                  <span className="text-sm font-semibold text-gray-900">Total: {formatCurrency(totalShown)}</span>
                </div>
                {loading ? (
                  <div className="p-12 text-center text-gray-400">Loading…</div>
                ) : filtered.length === 0 ? (
                  <div className="p-12 text-center">
                    <Receipt className="h-10 w-10 text-gray-200 mx-auto mb-3" />
                    <p className="text-gray-500">No expenses recorded yet</p>
                  </div>
                ) : (
                  <table className="w-full">
                    <thead>
                      <tr className="border-b border-gray-100">
                        <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Vendor</th>
                        <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Category</th>
                        <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Date</th>
                        <th className="text-right px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Amount</th>
                        <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Notes</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-50">
                      {filtered.map((e) => (
                        <tr key={e.id} className="hover:bg-gray-50/50 transition-colors">
                          <td className="px-6 py-3.5 text-sm font-medium text-gray-900">{e.vendor_name}</td>
                          <td className="px-6 py-3.5">
                            <span className="text-xs bg-gray-100 text-gray-600 px-2 py-0.5 rounded-full font-medium">
                              {e.category}
                            </span>
                          </td>
                          <td className="px-6 py-3.5 text-sm text-gray-500">{formatDate(e.date)}</td>
                          <td className="px-6 py-3.5 text-sm font-semibold text-gray-900 text-right">
                            {formatCurrency(e.amount)}
                          </td>
                          <td className="px-6 py-3.5 text-sm text-gray-500 max-w-xs truncate">{e.notes || '—'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            </div>

            {/* Sidebar: breakdown by category */}
            <div className="space-y-4">
              <div className="bg-white rounded-xl border border-gray-100 shadow-sm p-5">
                <p className="text-sm font-semibold text-gray-900 mb-1">All-time total</p>
                <p className="text-2xl font-bold text-gray-900">{formatCurrency(totalAll)}</p>
              </div>
              <div className="bg-white rounded-xl border border-gray-100 shadow-sm p-5">
                <p className="text-sm font-semibold text-gray-900 mb-4">By category</p>
                <div className="space-y-3">
                  {byCategory.length === 0 && <p className="text-sm text-gray-400">No data</p>}
                  {byCategory.map((c) => {
                    const pct = Math.round((c.total / totalAll) * 100)
                    return (
                      <div key={c.category}>
                        <div className="flex items-center justify-between mb-1">
                          <span className="text-xs text-gray-600">{c.category}</span>
                          <span className="text-xs font-medium text-gray-900">{formatCurrency(c.total)}</span>
                        </div>
                        <div className="w-full bg-gray-100 rounded-full h-1.5">
                          <div className="bg-blue-500 h-1.5 rounded-full" style={{ width: `${pct}%` }} />
                        </div>
                      </div>
                    )
                  })}
                </div>
              </div>
            </div>
          </div>
        </div>
      </main>

      <AddExpenseModal open={showAdd} onClose={() => setShowAdd(false)} onAdded={fetchExpenses} />
    </div>
  )
}

function AddExpenseModal({ open, onClose, onAdded }: { open: boolean; onClose: () => void; onAdded: () => void }) {
  const [form, setForm] = useState({ vendorName: '', category: 'Other', amount: '', date: '', notes: '' })
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState('')
  const supabase = createClient()

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setSaving(true)
    setError('')

    const { data: { user } } = await supabase.auth.getUser()

    const { error: err } = await supabase.from('expenses').insert({
      vendor_name: form.vendorName,
      category: form.category,
      amount: parseFloat(form.amount),
      date: form.date,
      notes: form.notes || null,
      created_by: user!.id,
    })

    if (err) { setError(err.message); setSaving(false); return }
    onAdded()
    onClose()
    setForm({ vendorName: '', category: 'Other', amount: '', date: '', notes: '' })
    setSaving(false)
  }

  return (
    <Modal open={open} onClose={onClose} title="Record Expense">
      <form onSubmit={handleSubmit} className="space-y-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">Vendor / Payee *</label>
          <input required type="text" value={form.vendorName} onChange={(e) => setForm({ ...form, vendorName: e.target.value })}
            className="w-full px-3.5 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            placeholder="ABC Landscaping" />
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Category *</label>
            <select required value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })}
              className="w-full px-3.5 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
              {CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Amount ($) *</label>
            <input required type="number" step="0.01" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })}
              className="w-full px-3.5 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
              placeholder="500.00" />
          </div>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">Date *</label>
          <input required type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })}
            className="w-full px-3.5 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">Notes</label>
          <textarea value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} rows={2}
            className="w-full px-3.5 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 resize-none"
            placeholder="Optional description…" />
        </div>
        {error && <p className="text-sm text-red-600">{error}</p>}
        <div className="flex justify-end gap-3 pt-2">
          <Button type="button" variant="outline" onClick={onClose}>Cancel</Button>
          <Button type="submit" loading={saving}>Save expense</Button>
        </div>
      </form>
    </Modal>
  )
}
