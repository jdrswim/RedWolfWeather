import { redirect } from 'next/navigation'
import { createClient } from '@/lib/supabaseServer'
import OwnerSidebar from '@/components/layout/OwnerSidebar'
import { StatCard } from '@/components/ui/Card'
import Badge, { statusBadge } from '@/components/ui/Badge'
import { formatCurrency, formatDate } from '@/lib/utils'
import { CreditCard, CheckCircle2, Clock, AlertTriangle } from 'lucide-react'
import Link from 'next/link'

export default async function OwnerDashboard() {
  const supabase = await createClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) redirect('/login')

  const [profileResult, settingsResult] = await Promise.all([
    supabase.from('profiles').select('*').eq('id', user.id).single(),
    supabase.from('hoa_settings').select('hoa_name').single(),
  ])

  const profile = profileResult.data
  if (profile?.role === 'admin') redirect('/dashboard/admin')

  const { data: dues } = await supabase
    .from('dues')
    .select('*')
    .eq('owner_id', user.id)
    .order('due_date', { ascending: false })
    .limit(6)

  const { data: payments } = await supabase
    .from('payments')
    .select('*')
    .eq('owner_id', user.id)
    .order('payment_date', { ascending: false })
    .limit(5)

  const outstanding = dues?.filter((d) => d.status !== 'paid').reduce((s: number, d: any) => s + d.balance_remaining, 0) ?? 0
  const totalPaid = payments?.reduce((s: number, p: any) => s + p.amount, 0) ?? 0
  const overdueCount = dues?.filter((d: any) => d.status === 'overdue').length ?? 0

  return (
    <div className="flex min-h-screen bg-gray-50">
      <OwnerSidebar
        userName={profile?.name}
        unitNumber={profile?.unit_number ?? undefined}
        hoaName={settingsResult.data?.hoa_name}
      />
      <main className="flex-1 ml-64 p-8">
        <div className="max-w-4xl mx-auto">
          <div className="mb-8">
            <h1 className="text-2xl font-bold text-gray-900">
              Welcome{profile?.name ? `, ${profile.name.split(' ')[0]}` : ''}
            </h1>
            <p className="text-gray-500 mt-1">
              {profile?.unit_number ? `Unit ${profile.unit_number}` : 'Owner'} ·{' '}
              {settingsResult.data?.hoa_name ?? 'Your HOA'}
            </p>
          </div>

          {/* Stats */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-5 mb-8">
            <StatCard
              title="Outstanding Balance"
              value={formatCurrency(outstanding)}
              icon={CreditCard}
              color={outstanding > 0 ? 'red' : 'green'}
            />
            <StatCard
              title="Total Paid"
              value={formatCurrency(totalPaid)}
              icon={CheckCircle2}
              color="green"
            />
            <StatCard
              title="Overdue"
              value={String(overdueCount)}
              subtitle={overdueCount > 0 ? 'Action needed' : 'All clear'}
              icon={overdueCount > 0 ? AlertTriangle : Clock}
              color={overdueCount > 0 ? 'red' : 'blue'}
            />
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {/* Dues */}
            <div className="bg-white rounded-xl border border-gray-100 shadow-sm">
              <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between">
                <h3 className="font-semibold text-gray-900">Recent Dues</h3>
                <Link href="/dashboard/owner/dues" className="text-xs text-blue-600 hover:text-blue-700 font-medium">
                  View all →
                </Link>
              </div>
              <div className="divide-y divide-gray-50">
                {!dues?.length ? (
                  <p className="px-6 py-8 text-center text-gray-400 text-sm">No dues records</p>
                ) : (
                  dues.map((due: any) => (
                    <div key={due.id} className="px-6 py-3.5 flex items-center justify-between">
                      <div>
                        <p className="text-sm font-medium text-gray-900">{due.month_year}</p>
                        <p className="text-xs text-gray-400">Due {formatDate(due.due_date)}</p>
                      </div>
                      <div className="text-right">
                        <p className="text-sm font-semibold text-gray-900">{formatCurrency(due.amount_due)}</p>
                        <Badge variant={statusBadge(due.status)} className="mt-1">
                          {due.status.charAt(0).toUpperCase() + due.status.slice(1)}
                        </Badge>
                      </div>
                    </div>
                  ))
                )}
              </div>
              {outstanding > 0 && (
                <div className="px-6 py-4 border-t border-gray-100">
                  <Link
                    href="/dashboard/owner/dues"
                    className="block w-full text-center bg-blue-600 text-white text-sm font-semibold py-2.5 rounded-lg hover:bg-blue-700 transition-colors"
                  >
                    Pay outstanding dues
                  </Link>
                </div>
              )}
            </div>

            {/* Recent payments */}
            <div className="bg-white rounded-xl border border-gray-100 shadow-sm">
              <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between">
                <h3 className="font-semibold text-gray-900">Payment History</h3>
              </div>
              <div className="divide-y divide-gray-50">
                {!payments?.length ? (
                  <p className="px-6 py-8 text-center text-gray-400 text-sm">No payments yet</p>
                ) : (
                  payments.map((p: any) => (
                    <div key={p.id} className="px-6 py-3.5 flex items-center justify-between">
                      <div>
                        <p className="text-sm font-medium text-gray-900">{formatCurrency(p.amount)}</p>
                        <p className="text-xs text-gray-400">{formatDate(p.payment_date)}</p>
                      </div>
                      <Badge variant={statusBadge(p.status)}>
                        {p.status.charAt(0).toUpperCase() + p.status.slice(1)}
                      </Badge>
                    </div>
                  ))
                )}
              </div>
            </div>
          </div>
        </div>
      </main>
    </div>
  )
}
