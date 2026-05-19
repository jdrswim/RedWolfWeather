'use client'

import { useEffect, useState, useCallback } from 'react'
import { createClient } from '@/lib/supabaseClient'
import OwnerSidebar from '@/components/layout/OwnerSidebar'
import PageHeader from '@/components/layout/PageHeader'
import Badge, { statusBadge } from '@/components/ui/Badge'
import Button from '@/components/ui/Button'
import { formatCurrency, formatDate } from '@/lib/utils'
import { CreditCard, CheckCircle2 } from 'lucide-react'
import type { Due, Profile, HoaSettings } from '@/types'

export default function OwnerDuesPage() {
  const [dues, setDues] = useState<Due[]>([])
  const [profile, setProfile] = useState<Profile | null>(null)
  const [settings, setSettings] = useState<HoaSettings | null>(null)
  const [loading, setLoading] = useState(true)
  const [payingId, setPayingId] = useState<string | null>(null)
  const supabase = createClient()

  const fetchData = useCallback(async () => {
    const { data: { user } } = await supabase.auth.getUser()
    if (!user) return

    const [{ data: p }, { data: d }, { data: s }] = await Promise.all([
      supabase.from('profiles').select('*').eq('id', user.id).single(),
      supabase.from('dues').select('*').eq('owner_id', user.id).order('due_date', { ascending: false }),
      supabase.from('hoa_settings').select('*').single(),
    ])

    setProfile(p)
    setDues(d ?? [])
    setSettings(s)
    setLoading(false)
  }, [supabase])

  useEffect(() => { fetchData() }, [fetchData])

  async function initiatePayment(due: Due) {
    setPayingId(due.id)
    try {
      const res = await fetch('/api/stripe/checkout', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ dueId: due.id, amount: due.balance_remaining }),
      })
      const { url, error } = await res.json()
      if (error) { alert(error); return }
      window.location.href = url
    } catch {
      alert('Failed to initiate payment. Please try again.')
    } finally {
      setPayingId(null)
    }
  }

  const outstanding = dues.filter((d) => d.status !== 'paid').reduce((s, d) => s + d.balance_remaining, 0)

  if (loading) {
    return (
      <div className="flex min-h-screen bg-gray-50">
        <OwnerSidebar hoaName={settings?.hoa_name} />
        <main className="flex-1 ml-64 p-8">
          <div className="text-center text-gray-400 py-20">Loading your dues…</div>
        </main>
      </div>
    )
  }

  return (
    <div className="flex min-h-screen bg-gray-50">
      <OwnerSidebar
        userName={profile?.name}
        unitNumber={profile?.unit_number ?? undefined}
        hoaName={settings?.hoa_name}
      />
      <main className="flex-1 ml-64 p-8">
        <div className="max-w-3xl mx-auto">
          <PageHeader
            title="My Dues"
            description="View and pay your HOA dues"
          />

          {/* Outstanding alert */}
          {outstanding > 0 && (
            <div className="bg-red-50 border border-red-100 rounded-xl p-5 mb-6">
              <div className="flex items-center justify-between">
                <div>
                  <p className="font-semibold text-red-900">Outstanding balance</p>
                  <p className="text-2xl font-bold text-red-600 mt-1">{formatCurrency(outstanding)}</p>
                </div>
                <CreditCard className="h-8 w-8 text-red-400" />
              </div>
            </div>
          )}

          {outstanding === 0 && dues.length > 0 && (
            <div className="bg-green-50 border border-green-100 rounded-xl p-4 mb-6 flex items-center gap-3">
              <CheckCircle2 className="h-5 w-5 text-green-500" />
              <p className="text-sm font-medium text-green-800">All dues are paid — you&apos;re up to date!</p>
            </div>
          )}

          {/* Dues list */}
          <div className="bg-white rounded-xl border border-gray-100 shadow-sm overflow-hidden">
            {dues.length === 0 ? (
              <div className="p-12 text-center text-gray-400">
                <CreditCard className="h-10 w-10 text-gray-200 mx-auto mb-3" />
                <p>No dues records yet</p>
              </div>
            ) : (
              <table className="w-full">
                <thead>
                  <tr className="border-b border-gray-100 bg-gray-50/50">
                    <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Period</th>
                    <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Amount</th>
                    <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Balance</th>
                    <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Due Date</th>
                    <th className="text-left px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Status</th>
                    <th className="text-right px-6 py-3 text-xs font-semibold text-gray-500 uppercase">Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {dues.map((due) => (
                    <tr key={due.id} className="hover:bg-gray-50/50 transition-colors">
                      <td className="px-6 py-4 text-sm font-medium text-gray-900">{due.month_year}</td>
                      <td className="px-6 py-4 text-sm text-gray-700">{formatCurrency(due.amount_due)}</td>
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
                          <Button
                            size="sm"
                            loading={payingId === due.id}
                            onClick={() => initiatePayment(due)}
                          >
                            <CreditCard className="h-3.5 w-3.5" />
                            Pay now
                          </Button>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>

          {dues.length > 0 && (
            <p className="text-xs text-gray-400 mt-4 text-center">
              Payments are processed securely via Stripe. Card and ACH available.
            </p>
          )}
        </div>
      </main>
    </div>
  )
}
