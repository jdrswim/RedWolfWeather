import { redirect } from 'next/navigation'
import { createClient } from '@/lib/supabaseServer'
import AdminSidebar from '@/components/layout/AdminSidebar'
import SetupChecklist from '@/components/onboarding/SetupChecklist'
import { StatCard } from '@/components/ui/Card'
import { formatCurrency, formatDate } from '@/lib/utils'
import { Users, CreditCard, Receipt, TrendingUp, AlertTriangle } from 'lucide-react'
import type { ChecklistItem } from '@/components/onboarding/SetupChecklist'

export default async function AdminDashboard() {
  const supabase = await createClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) redirect('/login')

  const { data: profile } = await supabase.from('profiles').select('role, name').eq('id', user.id).single()
  if (profile?.role !== 'admin') redirect('/dashboard/owner')

  const { data: settings } = await supabase.from('hoa_settings').select('*').single()

  // Stats queries run in parallel
  const [
    { count: ownerCount },
    { data: dues },
    { data: payments },
    { data: expenses },
    { data: recentPayments },
  ] = await Promise.all([
    supabase.from('profiles').select('*', { count: 'exact', head: true }).eq('role', 'owner'),
    supabase.from('dues').select('amount_due, balance_remaining, status'),
    supabase.from('payments').select('amount').eq('status', 'completed'),
    supabase.from('expenses').select('amount'),
    supabase
      .from('payments')
      .select('id, amount, payment_date, status, profiles(name, unit_number)')
      .eq('status', 'completed')
      .order('payment_date', { ascending: false })
      .limit(5),
  ])

  const totalDue = dues?.reduce((s, d) => s + (d.balance_remaining || 0), 0) ?? 0
  const totalCollected = payments?.reduce((s, p) => s + (p.amount || 0), 0) ?? 0
  const totalExpenses = expenses?.reduce((s, e) => s + (e.amount || 0), 0) ?? 0
  const overdueCount = dues?.filter((d) => d.status === 'overdue').length ?? 0

  const checklistItems: ChecklistItem[] = [
    {
      id: 'org',
      label: 'Configure HOA organization',
      description: 'Set your HOA name, address, and unit count',
      completed: !!(settings?.hoa_name),
      href: '/dashboard/admin/settings',
    },
    {
      id: 'stripe',
      label: 'Connect Stripe payments',
      description: 'Add your Stripe API keys to accept online payments',
      completed: settings?.stripe_configured ?? false,
      href: '/dashboard/admin/settings',
    },
    {
      id: 'owners',
      label: 'Add unit owners',
      description: 'Invite owners to join their accounts',
      completed: (ownerCount ?? 0) > 0,
      href: '/dashboard/admin/owners',
    },
    {
      id: 'dues',
      label: 'Set up dues',
      description: 'Configure monthly dues for your owners',
      completed: (dues?.length ?? 0) > 0,
      href: '/dashboard/admin/dues',
    },
    {
      id: 'documents',
      label: 'Upload HOA documents',
      description: 'Add bylaws, rules, and meeting minutes',
      completed: false,
      href: '/dashboard/admin/documents',
    },
  ]

  const allComplete = checklistItems.every((i) => i.completed)

  return (
    <div className="flex min-h-screen bg-gray-50">
      <AdminSidebar hoaName={settings?.hoa_name} />
      <main className="flex-1 ml-64 p-8">
        <div className="max-w-6xl mx-auto">
          <div className="mb-8">
            <h1 className="text-2xl font-bold text-gray-900">
              Good morning{profile?.name ? `, ${profile.name.split(' ')[0]}` : ''}
            </h1>
            <p className="text-gray-500 mt-1">
              {settings?.hoa_name ?? 'Your HOA'} — Admin Dashboard
            </p>
          </div>

          {/* Stats */}
          <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-5 mb-8">
            <StatCard
              title="Total Owners"
              value={String(ownerCount ?? 0)}
              icon={Users}
              color="blue"
            />
            <StatCard
              title="Collected (All Time)"
              value={formatCurrency(totalCollected)}
              icon={TrendingUp}
              color="green"
            />
            <StatCard
              title="Outstanding Balance"
              value={formatCurrency(totalDue)}
              icon={CreditCard}
              color={totalDue > 0 ? 'red' : 'green'}
            />
            <StatCard
              title="Total Expenses"
              value={formatCurrency(totalExpenses)}
              icon={Receipt}
              color="yellow"
            />
          </div>

          <div className="grid grid-cols-1 xl:grid-cols-3 gap-6">
            {/* Recent payments */}
            <div className="xl:col-span-2 bg-white rounded-xl border border-gray-100 shadow-sm">
              <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between">
                <h3 className="font-semibold text-gray-900">Recent Payments</h3>
                <a href="/dashboard/admin/dues" className="text-xs text-blue-600 hover:text-blue-700 font-medium">
                  View all →
                </a>
              </div>
              <div className="divide-y divide-gray-50">
                {recentPayments && recentPayments.length > 0 ? (
                  recentPayments.map((p: any) => (
                    <div key={p.id} className="px-6 py-3.5 flex items-center justify-between">
                      <div>
                        <p className="text-sm font-medium text-gray-900">
                          {p.profiles?.name ?? 'Unknown'}
                          {p.profiles?.unit_number && (
                            <span className="text-gray-400 font-normal"> · Unit {p.profiles.unit_number}</span>
                          )}
                        </p>
                        <p className="text-xs text-gray-400">{formatDate(p.payment_date)}</p>
                      </div>
                      <span className="text-sm font-semibold text-green-600">
                        +{formatCurrency(p.amount)}
                      </span>
                    </div>
                  ))
                ) : (
                  <div className="px-6 py-8 text-center text-gray-400 text-sm">
                    No payments yet
                  </div>
                )}
              </div>
            </div>

            {/* Sidebar: checklist + alerts */}
            <div className="space-y-5">
              {!allComplete && (
                <SetupChecklist items={checklistItems} />
              )}

              {overdueCount > 0 && (
                <div className="bg-red-50 border border-red-100 rounded-xl p-4">
                  <div className="flex items-center gap-2 mb-2">
                    <AlertTriangle className="h-4 w-4 text-red-500" />
                    <p className="text-sm font-semibold text-red-900">Overdue Dues</p>
                  </div>
                  <p className="text-sm text-red-700">
                    {overdueCount} unit{overdueCount > 1 ? 's have' : ' has'} overdue dues.
                  </p>
                  <a href="/dashboard/admin/dues" className="text-xs text-red-600 font-medium mt-2 inline-block hover:text-red-700">
                    Manage dues →
                  </a>
                </div>
              )}

              {allComplete && (
                <div className="bg-green-50 border border-green-100 rounded-xl p-4">
                  <p className="text-sm font-medium text-green-800">Your HOA portal is fully configured!</p>
                </div>
              )}
            </div>
          </div>
        </div>
      </main>
    </div>
  )
}
