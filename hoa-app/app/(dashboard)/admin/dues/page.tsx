'use client'

import { useEffect, useState, useCallback } from 'react'
import { createClient } from '@/lib/supabaseClient'
import AdminSidebar from '@/components/layout/AdminSidebar'
import PageHeader from '@/components/layout/PageHeader'
import Badge, { statusBadge } from '@/components/ui/Badge'
import Button from '@/components/ui/Button'
import Modal from '@/components/ui/Modal'
import { formatCurrency, formatDate } from '@/lib/utils'
import { Plus, Search } from 'lucide-react'
import type { Due, Profile } from '@/types'

type DueWithProfile = Due & { profiles: Profile }

export default function AdminDuesPage() {
  const [dues, setDues] = useState<DueWithProfile[]>([])
  const [filtered, setFiltered] = useState<DueWithProfile[]>([])
  const [search, setSearch] = useState('')
  const [statusFilter, setStatusFilter] = useState('all')
  const [loading, setLoading] = useState(true)
  const [showAdd, setShowAdd] = useState(false)
  const [owners, setOwners] = useState<Profile[]>([])
  const [hoaName, setHoaName] = useState('')
  const supabase = createClient()

  const fetchDues = useCallback(async () => {
    const { data } = await supabase
      .from('dues')
      .select('*, profiles(*)')
      .order('due_date', { ascending: false })
    setDues((data as DueWithProfile[]) ?? [])
    setLoading(false)
  }, [supabase])

  useEffect(() => {
    fetchDues()
    supabase.from('profiles').select('*').eq('role', 'owner').then(({ data }) => setOwners(data ?? []))
    supabase.from('hoa_settings').select('hoa_name').single().then(({ data }) => { if (data) setHoaName(data.hoa_name) })
  }, [fetchDues, supabase])

  useEffect(() => {
    let result = dues
    if (statusFilter !== 'all') result = result.filter((d) => d.status === statusFilter)
    if (search) {
      const q = search.toLowerCase()
      result = result.filter(
        (d) =>
          d.profiles?.name?.toLowerCase().includes(q) ||
          d.profiles?.unit_number?.toLowerCase().includes(q) ||
          d.month_year?.toLowerCase().includes(q)
      )
    }
    setFiltered(result)
  }, [dues, search, statusFilter])

  async function markPaid(dueId: string) {
    await supabase
      .from('dues')
      .update({ status: 'paid', balance_remaining: 0 })
      .eq('id', dueId)
    fetchDues()
  }

  const totals = {
    outstanding: dues.filter((d) => d.status !== 'paid').reduce((s, d) => s + d.balance_remaining, 0),
    collected: dues.filter((d) => d.status === 'paid').reduce((s, d) => s + d.amount_due, 0),
    overdue: dues.filter((d) => d.status === 'overdue').length,
  }

  return (
    <div className="flex min-h-screen bg-gray-50">
      <AdminSidebar hoaName={hoaName} />
      <main className="flex-1 ml-64 p-8">
        <div className="max-w-6xl mx-auto">
          <PageHeader
            title="Dues Management"
            description="Track and manage owner dues and balances"
            action={
              <Button onClick={() => setShowAdd(true)}>
                <Plus className="h-4 w-4" />
                Add dues
              </Button>
            }
          />

          {/* Summary cards */}
          <div className="grid grid-cols-3 gap-4 mb-6">
            {[
              { label: 'Outstanding', value: formatCurrency(totals.outstanding), color: 'text-red-600' },
              { label: 'Collected', value: formatCurrency(totals.collected), color: 'text-green-600' },
              { label: 'Overdue count', value: String(totals.overdue), color: 'text-yellow-600' },
            ].map((stat) => (
              <div key={stat.label} className="bg-white rounded-xl border border-gray-100 shadow-sm px-5 py-4">
                <p className="text-xs text-gray-500 font-medium">{stat.label}</p>
                <p className={`text-xl font-bold mt-1 ${stat.color}`}>{stat.value}</p>
              </div>
            ))}
          </div>

          {/* Filters */}
          <div className="flex gap-3 mb-5">
            <div className="relative flex-1">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-gray-400" />
              <input
                type="text"
                placeholder="Search owner or month…"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                className="w-full pl-9 pr-4 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"
              />
            </div>
            <select
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value)}
              className="px-3 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"
            >
              <option value="all">All statuses</option>
              <option value="unpaid">Unpaid</option>
              <option value="paid">Paid</option>
              <option value="overdue">Overdue</option>
              <option value="partial">Partial</option>
            </select>
          </div>

          {/* Table */}
          <div className="bg-white rounded-xl border border-gray-100 shadow-sm overflow-hidden">
            {loading ? (
              <div className="p-12 text-center text-gray-400">Loading dues…</div>
            ) : filtered.length === 0 ? (
              <div className="p-12 text-center text-gray-400">No dues found</div>
            ) : (
              <table className="w-full">
                <thead>
                  <tr className="border-b border-gray-100 bg-gray-50/50">
                    <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Owner</th>
                    <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Period</th>
                    <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Amount</th>
                    <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Balance</th>
                    <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Due Date</th>
                    <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Status</th>
                    <th className="text-right px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {filtered.map((due) => (
                    <tr key={due.id} className="hover:bg-gray-50/50 transition-colors">
                      <td className="px-6 py-4">
                        <p className="text-sm font-medium text-gray-900">{due.profiles?.name ?? '—'}</p>
                        <p className="text-xs text-gray-400">
                          {due.profiles?.unit_number ? `Unit ${due.profiles.unit_number}` : ''}
                        </p>
                      </td>
                      <td className="px-6 py-4 text-sm text-gray-700">{due.month_year}</td>
                      <td className="px-6 py-4 text-sm font-medium text-gray-900">{formatCurrency(due.amount_due)}</td>
                      <td className="px-6 py-4 text-sm font-medium text-red-600">
                        {due.balance_remaining > 0 ? formatCurrency(due.balance_remaining) : '—'}
                      </td>
                      <td className="px-6 py-4 text-sm text-gray-500">{formatDate(due.due_date)}</td>
                      <td className="px-6 py-4">
                        <Badge variant={statusBadge(due.status)}>
                          {due.status.charAt(0).toUpperCase() + due.status.slice(1)}
                        </Badge>
                      </td>
                      <td className="px-6 py-4 text-right">
                        {due.status !== 'paid' && (
                          <Button variant="outline" size="sm" onClick={() => markPaid(due.id)}>
                            Mark paid
                          </Button>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      </main>

      <AddDuesModal
        open={showAdd}
        onClose={() => setShowAdd(false)}
        onAdded={fetchDues}
        owners={owners}
      />
    </div>
  )
}

function AddDuesModal({
  open, onClose, onAdded, owners,
}: {
  open: boolean; onClose: () => void; onAdded: () => void; owners: Profile[]
}) {
  const [form, setForm] = useState({
    ownerId: '',
    amount: '',
    dueDate: '',
    monthYear: '',
    notes: '',
    bulk: false,
  })
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState('')
  const supabase = createClient()

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setSaving(true)
    setError('')

    const records = form.bulk
      ? owners.map((o) => ({
          owner_id: o.id,
          amount_due: parseFloat(form.amount),
          due_date: form.dueDate,
          month_year: form.monthYear,
          status: 'unpaid',
          balance_remaining: parseFloat(form.amount),
          notes: form.notes || null,
        }))
      : [
          {
            owner_id: form.ownerId,
            amount_due: parseFloat(form.amount),
            due_date: form.dueDate,
            month_year: form.monthYear,
            status: 'unpaid',
            balance_remaining: parseFloat(form.amount),
            notes: form.notes || null,
          },
        ]

    const { error: err } = await supabase.from('dues').insert(records)
    if (err) { setError(err.message); setSaving(false); return }

    onAdded()
    onClose()
    setForm({ ownerId: '', amount: '', dueDate: '', monthYear: '', notes: '', bulk: false })
    setSaving(false)
  }

  return (
    <Modal open={open} onClose={onClose} title="Create Dues">
      <form onSubmit={handleSubmit} className="space-y-4">
        <label className="flex items-center gap-3 p-3 rounded-lg border border-gray-100 cursor-pointer hover:bg-gray-50">
          <input
            type="checkbox"
            checked={form.bulk}
            onChange={(e) => setForm({ ...form, bulk: e.target.checked })}
            className="rounded border-gray-300 text-blue-600"
          />
          <div>
            <p className="text-sm font-medium text-gray-900">Bulk — create for all owners</p>
            <p className="text-xs text-gray-400">Creates a dues record for every owner at once</p>
          </div>
        </label>

        {!form.bulk && (
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Owner *</label>
            <select
              required={!form.bulk}
              value={form.ownerId}
              onChange={(e) => setForm({ ...form, ownerId: e.target.value })}
              className="w-full px-3.5 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            >
              <option value="">Select owner…</option>
              {owners.map((o) => (
                <option key={o.id} value={o.id}>
                  {o.name} {o.unit_number ? `(Unit ${o.unit_number})` : ''}
                </option>
              ))}
            </select>
          </div>
        )}

        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Amount ($) *</label>
            <input required type="number" step="0.01" value={form.amount}
              onChange={(e) => setForm({ ...form, amount: e.target.value })}
              className="w-full px-3.5 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
              placeholder="200.00" />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Due date *</label>
            <input required type="date" value={form.dueDate}
              onChange={(e) => setForm({ ...form, dueDate: e.target.value })}
              className="w-full px-3.5 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
          </div>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">Month / Period *</label>
          <input required type="text" value={form.monthYear}
            onChange={(e) => setForm({ ...form, monthYear: e.target.value })}
            className="w-full px-3.5 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            placeholder="January 2025" />
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">Notes</label>
          <textarea value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })}
            rows={2}
            className="w-full px-3.5 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 resize-none"
            placeholder="Optional notes…" />
        </div>

        {error && <p className="text-sm text-red-600">{error}</p>}
        <div className="flex justify-end gap-3 pt-2">
          <Button type="button" variant="outline" onClick={onClose}>Cancel</Button>
          <Button type="submit" loading={saving}>Create dues</Button>
        </div>
      </form>
    </Modal>
  )
}
