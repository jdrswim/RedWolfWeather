'use client'

import { useState } from 'react'
import { createClient } from '@/lib/supabaseClient'
import { Plus } from 'lucide-react'
import { useRouter } from 'next/navigation'

const CATEGORIES = ['Insurance', 'Utilities', 'Landscaping', 'Repairs', 'Management', 'Legal', 'Accounting', 'Other']

export default function AddExpenseForm() {
  const [open, setOpen] = useState(false)
  const [vendor, setVendor] = useState('')
  const [category, setCategory] = useState(CATEGORIES[0])
  const [amount, setAmount] = useState('')
  const [date, setDate] = useState(new Date().toISOString().split('T')[0])
  const [notes, setNotes] = useState('')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const router = useRouter()
  const supabase = createClient()

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setLoading(true)
    setError('')

    const { error: insertError } = await supabase.from('expenses').insert({
      vendor_name: vendor,
      category,
      amount: parseFloat(amount),
      date,
      notes,
    })

    if (insertError) { setError(insertError.message); setLoading(false); return }

    setOpen(false)
    setVendor(''); setAmount(''); setNotes('')
    router.refresh()
    setLoading(false)
  }

  return (
    <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6">
      <div className="flex items-center justify-between">
        <h2 className="font-semibold text-gray-900">Add Expense</h2>
        <button onClick={() => setOpen(!open)}
          className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold rounded-xl transition-colors">
          <Plus size={16} />
          {open ? 'Cancel' : 'Add Expense'}
        </button>
      </div>

      {open && (
        <form onSubmit={handleSubmit} className="mt-5 grid grid-cols-1 sm:grid-cols-2 gap-4">
          {error && <div className="sm:col-span-2 p-3 bg-red-50 border border-red-200 rounded-lg text-red-700 text-sm">{error}</div>}
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Vendor Name</label>
            <input type="text" value={vendor} onChange={e => setVendor(e.target.value)} required
              className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500"
              placeholder="ABC Landscaping" />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Category</label>
            <select value={category} onChange={e => setCategory(e.target.value)}
              className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white">
              {CATEGORIES.map(c => <option key={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Amount ($)</label>
            <input type="number" step="0.01" min="0" value={amount} onChange={e => setAmount(e.target.value)} required
              className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500"
              placeholder="500.00" />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Date</label>
            <input type="date" value={date} onChange={e => setDate(e.target.value)} required
              className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500" />
          </div>
          <div className="sm:col-span-2">
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Notes (optional)</label>
            <textarea value={notes} onChange={e => setNotes(e.target.value)} rows={2}
              className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 resize-none"
              placeholder="Monthly lawn maintenance..." />
          </div>
          <div className="sm:col-span-2">
            <button type="submit" disabled={loading}
              className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-400 text-white text-sm font-semibold rounded-xl transition-colors">
              {loading ? 'Saving...' : 'Save Expense'}
            </button>
          </div>
        </form>
      )}
    </div>
  )
}
