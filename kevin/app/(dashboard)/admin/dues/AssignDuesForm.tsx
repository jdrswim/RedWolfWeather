'use client'

import { useState } from 'react'
import { createClient } from '@/lib/supabaseClient'
import { Plus } from 'lucide-react'
import { useRouter } from 'next/navigation'

interface Owner { id: string; name: string; unit_number: string | null }

export default function AssignDuesForm({ owners }: { owners: Owner[] }) {
  const [open, setOpen] = useState(false)
  const [ownerId, setOwnerId] = useState('')
  const [amount, setAmount] = useState('')
  const [dueDate, setDueDate] = useState('')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const router = useRouter()
  const supabase = createClient()

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setLoading(true)
    setError('')

    const { error: insertError } = await supabase.from('dues').insert({
      owner_id: ownerId,
      amount_due: parseFloat(amount),
      balance_remaining: parseFloat(amount),
      due_date: dueDate,
      status: 'pending',
    })

    if (insertError) {
      setError(insertError.message)
      setLoading(false)
      return
    }

    setOpen(false)
    setOwnerId('')
    setAmount('')
    setDueDate('')
    router.refresh()
    setLoading(false)
  }

  return (
    <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6">
      <div className="flex items-center justify-between">
        <h2 className="font-semibold text-gray-900">Assign New Dues</h2>
        <button
          onClick={() => setOpen(!open)}
          className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold rounded-xl transition-colors"
        >
          <Plus size={16} />
          {open ? 'Cancel' : 'Assign Dues'}
        </button>
      </div>

      {open && (
        <form onSubmit={handleSubmit} className="mt-5 grid grid-cols-1 sm:grid-cols-3 gap-4">
          {error && <div className="sm:col-span-3 p-3 bg-red-50 border border-red-200 rounded-lg text-red-700 text-sm">{error}</div>}
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Owner</label>
            <select
              value={ownerId} onChange={e => setOwnerId(e.target.value)} required
              className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"
            >
              <option value="">Select owner...</option>
              {owners.map(o => (
                <option key={o.id} value={o.id}>
                  {o.name} {o.unit_number ? `(Unit ${o.unit_number})` : ''}
                </option>
              ))}
            </select>
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Amount ($)</label>
            <input
              type="number" step="0.01" min="0" value={amount} onChange={e => setAmount(e.target.value)} required
              className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500"
              placeholder="250.00"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Due Date</label>
            <input
              type="date" value={dueDate} onChange={e => setDueDate(e.target.value)} required
              className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>
          <div className="sm:col-span-3">
            <button type="submit" disabled={loading}
              className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-400 text-white text-sm font-semibold rounded-xl transition-colors">
              {loading ? 'Assigning...' : 'Assign Dues'}
            </button>
          </div>
        </form>
      )}
    </div>
  )
}
