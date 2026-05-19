'use client'

import { useState } from 'react'
import { createClient } from '@/lib/supabaseClient'
import { UserPlus, X } from 'lucide-react'
import { useRouter } from 'next/navigation'

export default function AddOwnerModal() {
  const [open, setOpen] = useState(false)
  const [name, setName] = useState('')
  const [email, setEmail] = useState('')
  const [unit, setUnit] = useState('')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const router = useRouter()

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setLoading(true)
    setError('')

    const res = await fetch('/kevin/api/admin/create-owner', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, email, unitNumber: unit }),
    })

    const data = await res.json()
    if (!res.ok) {
      setError(data.error || 'Failed to create owner')
      setLoading(false)
      return
    }

    setOpen(false)
    setName('')
    setEmail('')
    setUnit('')
    router.refresh()
    setLoading(false)
  }

  return (
    <>
      <button
        onClick={() => setOpen(true)}
        className="flex items-center gap-2 px-4 py-2.5 bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold rounded-xl transition-colors"
      >
        <UserPlus size={16} />
        Add Owner
      </button>

      {open && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md p-6">
            <div className="flex items-center justify-between mb-6">
              <h2 className="text-lg font-semibold text-gray-900">Add New Owner</h2>
              <button onClick={() => setOpen(false)} className="text-gray-400 hover:text-gray-600">
                <X size={20} />
              </button>
            </div>

            {error && (
              <div className="mb-4 p-3 bg-red-50 border border-red-200 rounded-lg text-red-700 text-sm">{error}</div>
            )}

            <form onSubmit={handleSubmit} className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1.5">Full Name</label>
                <input
                  type="text" value={name} onChange={e => setName(e.target.value)} required
                  className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500"
                  placeholder="Jane Smith"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1.5">Email Address</label>
                <input
                  type="email" value={email} onChange={e => setEmail(e.target.value)} required
                  className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500"
                  placeholder="jane@example.com"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1.5">Unit Number</label>
                <input
                  type="text" value={unit} onChange={e => setUnit(e.target.value)}
                  className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500"
                  placeholder="e.g. 101"
                />
              </div>
              <div className="flex gap-3 pt-2">
                <button type="button" onClick={() => setOpen(false)}
                  className="flex-1 py-2.5 border border-gray-300 rounded-xl text-sm font-medium text-gray-700 hover:bg-gray-50 transition-colors">
                  Cancel
                </button>
                <button type="submit" disabled={loading}
                  className="flex-1 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-400 rounded-xl text-sm font-semibold text-white transition-colors">
                  {loading ? 'Creating...' : 'Create Owner'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </>
  )
}
