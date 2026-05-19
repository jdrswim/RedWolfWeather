import { createServerSupabaseClient } from '@/lib/supabaseServer'
import { redirect } from 'next/navigation'
import { Users, Mail, Home, CheckCircle, XCircle } from 'lucide-react'
import { format } from 'date-fns'
import AddOwnerModal from './AddOwnerModal'

export default async function AdminOwnersPage() {
  const supabase = await createServerSupabaseClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) redirect('/kevin/login')

  const { data: profile } = await supabase.from('profiles').select('role').eq('id', user.id).single()
  if (profile?.role !== 'admin') redirect('/kevin/owner')

  const { data: owners } = await supabase
    .from('profiles')
    .select('*, dues(balance_remaining, status)')
    .eq('role', 'owner')
    .order('name')

  return (
    <div className="max-w-6xl mx-auto space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-gray-900">Owners</h1>
          <p className="text-gray-500 text-sm mt-1">{owners?.length || 0} homeowners registered</p>
        </div>
        <AddOwnerModal />
      </div>

      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
        {!owners || owners.length === 0 ? (
          <div className="text-center py-16 text-gray-400">
            <Users className="w-12 h-12 mx-auto mb-3 opacity-40" />
            <p className="font-medium">No owners yet</p>
            <p className="text-sm mt-1">Add your first homeowner to get started</p>
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="bg-gray-50 border-b border-gray-100">
                <tr>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Owner</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Unit</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Balance Due</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Status</th>
                  <th className="text-left px-6 py-4 font-medium text-gray-500">Joined</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                {owners.map(owner => {
                  const balance = (owner.dues as any[])?.reduce((s: number, d: any) => s + (d.balance_remaining || 0), 0) || 0
                  return (
                    <tr key={owner.id} className="hover:bg-gray-50 transition-colors">
                      <td className="px-6 py-4">
                        <div className="flex items-center gap-3">
                          <div className="w-8 h-8 bg-blue-100 rounded-full flex items-center justify-center text-blue-700 font-bold text-xs">
                            {owner.name?.charAt(0)?.toUpperCase() || '?'}
                          </div>
                          <div>
                            <p className="font-medium text-gray-900">{owner.name || 'Unknown'}</p>
                            <p className="text-gray-400 text-xs">{owner.email}</p>
                          </div>
                        </div>
                      </td>
                      <td className="px-6 py-4 text-gray-600">{owner.unit_number || '—'}</td>
                      <td className="px-6 py-4">
                        <span className={balance > 0 ? 'text-red-600 font-semibold' : 'text-green-600 font-semibold'}>
                          ${balance.toFixed(2)}
                        </span>
                      </td>
                      <td className="px-6 py-4">
                        <div className="flex items-center gap-1.5">
                          {owner.is_active ? (
                            <><CheckCircle className="w-4 h-4 text-green-500" /><span className="text-green-700 text-xs font-medium">Active</span></>
                          ) : (
                            <><XCircle className="w-4 h-4 text-gray-400" /><span className="text-gray-500 text-xs font-medium">Inactive</span></>
                          )}
                        </div>
                      </td>
                      <td className="px-6 py-4 text-gray-500">
                        {owner.created_at ? format(new Date(owner.created_at), 'MMM d, yyyy') : '—'}
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  )
}
