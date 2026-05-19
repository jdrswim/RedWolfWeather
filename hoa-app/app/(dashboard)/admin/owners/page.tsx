'use client'

import { useEffect, useState, useCallback } from 'react'
import { createClient } from '@/lib/supabaseClient'
import AdminSidebar from '@/components/layout/AdminSidebar'
import PageHeader from '@/components/layout/PageHeader'
import Badge, { statusBadge } from '@/components/ui/Badge'
import Button from '@/components/ui/Button'
import Modal from '@/components/ui/Modal'
import { formatDate, getInitials } from '@/lib/utils'
import { UserPlus, Search, Mail, Phone, Home } from 'lucide-react'
import type { Profile } from '@/types'

export default function OwnersPage() {
  const [owners, setOwners] = useState<Profile[]>([])
  const [filtered, setFiltered] = useState<Profile[]>([])
  const [search, setSearch] = useState('')
  const [loading, setLoading] = useState(true)
  const [showAddModal, setShowAddModal] = useState(false)
  const [selectedOwner, setSelectedOwner] = useState<Profile | null>(null)
  const [hoaName, setHoaName] = useState('')

  const supabase = createClient()

  const fetchOwners = useCallback(async () => {
    const { data } = await supabase
      .from('profiles')
      .select('*')
      .eq('role', 'owner')
      .order('unit_number')
    setOwners(data ?? [])
    setFiltered(data ?? [])
    setLoading(false)
  }, [supabase])

  useEffect(() => {
    fetchOwners()
    supabase.from('hoa_settings').select('hoa_name').single().then(({ data }) => {
      if (data) setHoaName(data.hoa_name)
    })
  }, [fetchOwners, supabase])

  useEffect(() => {
    if (!search) {
      setFiltered(owners)
    } else {
      const q = search.toLowerCase()
      setFiltered(
        owners.filter(
          (o) =>
            o.name?.toLowerCase().includes(q) ||
            o.email?.toLowerCase().includes(q) ||
            o.unit_number?.toLowerCase().includes(q)
        )
      )
    }
  }, [search, owners])

  return (
    <div className="flex min-h-screen bg-gray-50">
      <AdminSidebar hoaName={hoaName} />
      <main className="flex-1 ml-64 p-8">
        <div className="max-w-5xl mx-auto">
          <PageHeader
            title="Owners"
            description="Manage unit owners and their accounts"
            action={
              <Button onClick={() => setShowAddModal(true)}>
                <UserPlus className="h-4 w-4" />
                Add owner
              </Button>
            }
          />

          {/* Search */}
          <div className="relative mb-6">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-gray-400" />
            <input
              type="text"
              placeholder="Search by name, email, or unit…"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="w-full pl-9 pr-4 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"
            />
          </div>

          {/* Table */}
          <div className="bg-white rounded-xl border border-gray-100 shadow-sm overflow-hidden">
            {loading ? (
              <div className="p-12 text-center text-gray-400">Loading owners…</div>
            ) : filtered.length === 0 ? (
              <div className="p-12 text-center">
                <Home className="h-10 w-10 text-gray-200 mx-auto mb-3" />
                <p className="text-gray-500 font-medium">No owners found</p>
                <p className="text-gray-400 text-sm mt-1">Add your first owner to get started</p>
              </div>
            ) : (
              <table className="w-full">
                <thead>
                  <tr className="border-b border-gray-100 bg-gray-50/50">
                    <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wider">Owner</th>
                    <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wider">Unit</th>
                    <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wider">Contact</th>
                    <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wider">Joined</th>
                    <th className="text-right px-6 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wider">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {filtered.map((owner) => (
                    <tr key={owner.id} className="hover:bg-gray-50/50 transition-colors">
                      <td className="px-6 py-4">
                        <div className="flex items-center gap-3">
                          <div className="h-9 w-9 bg-blue-100 rounded-full flex items-center justify-center flex-shrink-0">
                            <span className="text-xs font-semibold text-blue-700">
                              {getInitials(owner.name || owner.email || '?')}
                            </span>
                          </div>
                          <div>
                            <p className="text-sm font-medium text-gray-900">{owner.name || 'No name'}</p>
                            <p className="text-xs text-gray-400">{owner.email}</p>
                          </div>
                        </div>
                      </td>
                      <td className="px-6 py-4">
                        <span className="text-sm text-gray-700 font-medium">
                          {owner.unit_number ? `Unit ${owner.unit_number}` : '—'}
                        </span>
                      </td>
                      <td className="px-6 py-4">
                        <div className="flex flex-col gap-1">
                          <div className="flex items-center gap-1.5 text-xs text-gray-500">
                            <Mail className="h-3 w-3" />
                            {owner.email}
                          </div>
                          {owner.phone && (
                            <div className="flex items-center gap-1.5 text-xs text-gray-500">
                              <Phone className="h-3 w-3" />
                              {owner.phone}
                            </div>
                          )}
                        </div>
                      </td>
                      <td className="px-6 py-4 text-sm text-gray-500">
                        {formatDate(owner.created_at)}
                      </td>
                      <td className="px-6 py-4 text-right">
                        <Button
                          variant="ghost"
                          size="sm"
                          onClick={() => setSelectedOwner(owner)}
                        >
                          View
                        </Button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      </main>

      <AddOwnerModal
        open={showAddModal}
        onClose={() => setShowAddModal(false)}
        onAdded={fetchOwners}
      />

      {selectedOwner && (
        <OwnerDetailModal
          owner={selectedOwner}
          onClose={() => setSelectedOwner(null)}
        />
      )}
    </div>
  )
}

function AddOwnerModal({ open, onClose, onAdded }: { open: boolean; onClose: () => void; onAdded: () => void }) {
  const [form, setForm] = useState({ name: '', email: '', unitNumber: '', phone: '' })
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState('')
  const supabase = createClient()

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setSaving(true)
    setError('')

    // Create auth user via admin (service role needed for invite flow; here we insert profile directly)
    // In production, use Supabase Admin API to invite user by email
    const { error: err } = await supabase.from('profiles').insert({
      name: form.name,
      email: form.email,
      role: 'owner',
      unit_number: form.unitNumber || null,
      phone: form.phone || null,
    })

    if (err) {
      setError(err.message)
    } else {
      onAdded()
      onClose()
      setForm({ name: '', email: '', unitNumber: '', phone: '' })
    }
    setSaving(false)
  }

  return (
    <Modal open={open} onClose={onClose} title="Add Owner">
      <form onSubmit={handleSubmit} className="space-y-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">Full name *</label>
          <input required type="text" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })}
            className="w-full px-3.5 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            placeholder="Jane Smith" />
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">Email *</label>
          <input required type="email" value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })}
            className="w-full px-3.5 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            placeholder="jane@example.com" />
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Unit number</label>
            <input type="text" value={form.unitNumber} onChange={(e) => setForm({ ...form, unitNumber: e.target.value })}
              className="w-full px-3.5 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
              placeholder="4B" />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Phone</label>
            <input type="tel" value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })}
              className="w-full px-3.5 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
              placeholder="555-0100" />
          </div>
        </div>
        {error && <p className="text-sm text-red-600">{error}</p>}
        <div className="flex justify-end gap-3 pt-2">
          <Button type="button" variant="outline" onClick={onClose}>Cancel</Button>
          <Button type="submit" loading={saving}>Add owner</Button>
        </div>
      </form>
    </Modal>
  )
}

function OwnerDetailModal({ owner, onClose }: { owner: Profile; onClose: () => void }) {
  return (
    <Modal open onClose={onClose} title="Owner Details">
      <div className="space-y-4">
        <div className="flex items-center gap-4">
          <div className="h-14 w-14 bg-blue-100 rounded-full flex items-center justify-center">
            <span className="text-lg font-semibold text-blue-700">
              {getInitials(owner.name || owner.email || '?')}
            </span>
          </div>
          <div>
            <p className="font-semibold text-gray-900">{owner.name}</p>
            <p className="text-sm text-gray-500">{owner.email}</p>
          </div>
        </div>
        <div className="grid grid-cols-2 gap-4 pt-2">
          <div>
            <p className="text-xs text-gray-400 font-medium">Unit</p>
            <p className="text-sm text-gray-900 mt-1">{owner.unit_number ? `Unit ${owner.unit_number}` : '—'}</p>
          </div>
          <div>
            <p className="text-xs text-gray-400 font-medium">Phone</p>
            <p className="text-sm text-gray-900 mt-1">{owner.phone || '—'}</p>
          </div>
          <div>
            <p className="text-xs text-gray-400 font-medium">Role</p>
            <Badge variant="info" className="mt-1">Owner</Badge>
          </div>
          <div>
            <p className="text-xs text-gray-400 font-medium">Joined</p>
            <p className="text-sm text-gray-900 mt-1">{formatDate(owner.created_at)}</p>
          </div>
        </div>
      </div>
    </Modal>
  )
}
