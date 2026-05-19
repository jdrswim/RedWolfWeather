'use client'

import { useState } from 'react'
import { createClient } from '@/lib/supabaseClient'
import { useRouter } from 'next/navigation'
import { CheckCircle2, Circle, Building2, CreditCard, Settings, User, FileText, ExternalLink } from 'lucide-react'
import clsx from 'clsx'

interface ChecklistItem {
  step_key: string
  step_label: string
  completed: boolean
  sort_order: number
}

interface Props {
  settings: any
  checklist: ChecklistItem[]
}

const steps = [
  { key: 'hoa_info', label: 'HOA Organization', icon: Building2 },
  { key: 'admin_identity', label: 'Admin Identity', icon: User },
  { key: 'financial_setup', label: 'Financial Setup', icon: CreditCard },
  { key: 'payment_config', label: 'Payment Config', icon: Settings },
  { key: 'documents_upload', label: 'Documents', icon: FileText },
]

export default function SettingsForm({ settings: initialSettings, checklist }: Props) {
  const [activeTab, setActiveTab] = useState('hoa_info')
  const [saving, setSaving] = useState(false)
  const [savedMsg, setSavedMsg] = useState('')
  const router = useRouter()
  const supabase = createClient()

  const [hoaName, setHoaName] = useState(initialSettings?.hoa_name || '')
  const [address, setAddress] = useState(initialSettings?.address || '')
  const [numUnits, setNumUnits] = useState(initialSettings?.num_units?.toString() || '')
  const [defaultDues, setDefaultDues] = useState(initialSettings?.default_monthly_dues?.toString() || '')
  const [defaultDueDay, setDefaultDueDay] = useState(initialSettings?.default_due_day?.toString() || '1')
  const [lateFee, setLateFee] = useState(initialSettings?.late_fee_amount?.toString() || '0')
  const [lateGraceDays, setLateGraceDays] = useState(initialSettings?.late_fee_grace_days?.toString() || '5')
  const [accountingStart, setAccountingStart] = useState(initialSettings?.accounting_start_date || '')
  const [paymentMethods, setPaymentMethods] = useState<string[]>(initialSettings?.payment_methods || ['card'])

  function isStepDone(key: string) {
    return checklist.find(c => c.step_key === key)?.completed || false
  }

  async function markStepComplete(stepKey: string) {
    await supabase.from('setup_checklist')
      .update({ completed: true, completed_at: new Date().toISOString() })
      .eq('step_key', stepKey)
  }

  async function saveHoaInfo() {
    setSaving(true)
    const { error } = await supabase.from('hoa_settings')
      .update({
        hoa_name: hoaName,
        address,
        num_units: parseInt(numUnits) || 0,
        default_monthly_dues: parseFloat(defaultDues) || 0,
      })
      .eq('id', '00000000-0000-0000-0000-000000000001')
    if (!error) {
      await markStepComplete('hoa_info')
      setSavedMsg('HOA settings saved!')
      router.refresh()
    }
    setSaving(false)
    setTimeout(() => setSavedMsg(''), 3000)
  }

  async function savePaymentConfig() {
    setSaving(true)
    const { error } = await supabase.from('hoa_settings')
      .update({
        default_due_day: parseInt(defaultDueDay) || 1,
        late_fee_amount: parseFloat(lateFee) || 0,
        late_fee_grace_days: parseInt(lateGraceDays) || 5,
        payment_methods: paymentMethods,
      })
      .eq('id', '00000000-0000-0000-0000-000000000001')
    if (!error) {
      await markStepComplete('payment_config')
      setSavedMsg('Payment config saved!')
      router.refresh()
    }
    setSaving(false)
    setTimeout(() => setSavedMsg(''), 3000)
  }

  async function saveFinancialSetup() {
    setSaving(true)
    const { error } = await supabase.from('hoa_settings')
      .update({ accounting_start_date: accountingStart || null })
      .eq('id', '00000000-0000-0000-0000-000000000001')
    if (!error) {
      await markStepComplete('financial_setup')
      setSavedMsg('Financial settings saved!')
      router.refresh()
    }
    setSaving(false)
    setTimeout(() => setSavedMsg(''), 3000)
  }

  async function markAdminIdentityDone() {
    await markStepComplete('admin_identity')
    setSavedMsg('Admin identity confirmed!')
    router.refresh()
    setTimeout(() => setSavedMsg(''), 3000)
  }

  async function markDocumentsDone() {
    await markStepComplete('documents_upload')
    setSavedMsg('Documents step complete!')
    router.refresh()
    setTimeout(() => setSavedMsg(''), 3000)
  }

  const totalDone = checklist.filter(c => c.completed).length
  const total = checklist.length

  return (
    <div className="space-y-6">
      {/* Progress bar */}
      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6">
        <div className="flex items-center justify-between mb-3">
          <h2 className="font-semibold text-gray-900">Setup Progress</h2>
          <span className="text-sm text-gray-500 font-medium">{totalDone}/{total} complete</span>
        </div>
        <div className="w-full h-2.5 bg-gray-100 rounded-full overflow-hidden">
          <div
            className="h-full bg-blue-600 rounded-full transition-all duration-500"
            style={{ width: `${(totalDone / total) * 100}%` }}
          />
        </div>
        {totalDone === total && (
          <p className="text-green-600 text-sm font-medium mt-3 flex items-center gap-1.5">
            <CheckCircle2 size={16} /> Setup complete! Your HOA portal is fully configured.
          </p>
        )}
      </div>

      {savedMsg && (
        <div className="p-3 bg-green-50 border border-green-200 rounded-xl text-green-700 text-sm font-medium">
          {savedMsg}
        </div>
      )}

      {/* Step tabs */}
      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="flex border-b border-gray-100 overflow-x-auto">
          {steps.map(({ key, label, icon: Icon }) => {
            const done = isStepDone(key)
            return (
              <button
                key={key}
                onClick={() => setActiveTab(key)}
                className={clsx(
                  'flex items-center gap-2 px-5 py-4 text-sm font-medium whitespace-nowrap transition-colors border-b-2 -mb-px',
                  activeTab === key
                    ? 'border-blue-600 text-blue-600 bg-blue-50'
                    : 'border-transparent text-gray-500 hover:text-gray-700 hover:bg-gray-50'
                )}
              >
                {done ? <CheckCircle2 size={16} className="text-green-500" /> : <Circle size={16} className="text-gray-300" />}
                {label}
              </button>
            )
          })}
        </div>

        <div className="p-6">
          {/* HOA Info */}
          {activeTab === 'hoa_info' && (
            <div className="space-y-5">
              <h3 className="text-lg font-semibold text-gray-900">HOA Organization Settings</h3>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1.5">HOA Name</label>
                  <input type="text" value={hoaName} onChange={e => setHoaName(e.target.value)}
                    className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500"
                    placeholder="Sunset Hills HOA" />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1.5">Number of Units</label>
                  <input type="number" min="1" value={numUnits} onChange={e => setNumUnits(e.target.value)}
                    className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500"
                    placeholder="24" />
                </div>
                <div className="sm:col-span-2">
                  <label className="block text-sm font-medium text-gray-700 mb-1.5">Address</label>
                  <input type="text" value={address} onChange={e => setAddress(e.target.value)}
                    className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500"
                    placeholder="123 Main Street, City, State 12345" />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1.5">Default Monthly Dues ($)</label>
                  <input type="number" step="0.01" min="0" value={defaultDues} onChange={e => setDefaultDues(e.target.value)}
                    className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500"
                    placeholder="250.00" />
                </div>
              </div>
              <button onClick={saveHoaInfo} disabled={saving}
                className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-400 text-white text-sm font-semibold rounded-xl transition-colors">
                {saving ? 'Saving...' : 'Save HOA Settings'}
              </button>
            </div>
          )}

          {/* Admin Identity */}
          {activeTab === 'admin_identity' && (
            <div className="space-y-5">
              <h3 className="text-lg font-semibold text-gray-900">Admin Identity</h3>
              <div className="p-4 bg-blue-50 rounded-xl border border-blue-100">
                <p className="text-sm text-blue-800 font-medium mb-2">Your admin account is already configured.</p>
                <p className="text-sm text-blue-700">
                  You signed up with your email and created a password. To change your admin role or add additional admins, update the <code className="bg-blue-100 px-1 rounded">profiles</code> table in Supabase.
                </p>
              </div>
              <div className="p-4 bg-amber-50 rounded-xl border border-amber-100">
                <p className="text-sm font-semibold text-amber-800 mb-1">To make another user an admin:</p>
                <code className="text-xs bg-amber-100 px-3 py-2 rounded-lg block text-amber-900 font-mono">
                  UPDATE profiles SET role = &apos;admin&apos; WHERE email = &apos;admin@example.com&apos;;
                </code>
                <p className="text-xs text-amber-600 mt-2">Run this in your Supabase SQL editor.</p>
              </div>
              <button onClick={markAdminIdentityDone}
                className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold rounded-xl transition-colors">
                Mark as Complete
              </button>
            </div>
          )}

          {/* Financial Setup */}
          {activeTab === 'financial_setup' && (
            <div className="space-y-5">
              <h3 className="text-lg font-semibold text-gray-900">Financial Setup</h3>

              <div className="p-4 bg-amber-50 rounded-xl border border-amber-100">
                <p className="text-sm font-semibold text-amber-800 mb-2">Stripe Configuration</p>
                <p className="text-sm text-amber-700 mb-3">
                  Stripe API keys are configured via environment variables — NOT stored in the database.
                  Your Stripe secret key and webhook secret must be set in your Vercel/hosting environment.
                </p>
                <ul className="text-xs text-amber-700 space-y-1 font-mono">
                  <li>STRIPE_SECRET_KEY=sk_live_...</li>
                  <li>NEXT_PUBLIC_STRIPE_PUBLIC_KEY=pk_live_...</li>
                  <li>STRIPE_WEBHOOK_SECRET=whsec_...</li>
                </ul>
                <a href="https://dashboard.stripe.com/apikeys" target="_blank" rel="noopener noreferrer"
                  className="mt-3 inline-flex items-center gap-1.5 text-xs font-medium text-amber-800 hover:text-amber-900">
                  <ExternalLink size={12} /> Open Stripe Dashboard
                </a>
              </div>

              <div className="p-4 bg-blue-50 rounded-xl border border-blue-100">
                <p className="text-sm font-semibold text-blue-800 mb-2">Bank Payout Account</p>
                <p className="text-sm text-blue-700">
                  Bank account connection is handled directly in your Stripe Dashboard under{' '}
                  <strong>Settings → Bank Accounts</strong>. This is not configured in the app — Stripe handles all payouts.
                </p>
                <a href="https://dashboard.stripe.com/settings/payouts" target="_blank" rel="noopener noreferrer"
                  className="mt-2 inline-flex items-center gap-1.5 text-xs font-medium text-blue-800 hover:text-blue-900">
                  <ExternalLink size={12} /> Configure Payouts in Stripe
                </a>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1.5">Accounting Start Date</label>
                <input type="date" value={accountingStart} onChange={e => setAccountingStart(e.target.value)}
                  className="w-full sm:w-64 px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500" />
              </div>

              <button onClick={saveFinancialSetup} disabled={saving}
                className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-400 text-white text-sm font-semibold rounded-xl transition-colors">
                {saving ? 'Saving...' : 'Save Financial Settings'}
              </button>
            </div>
          )}

          {/* Payment Config */}
          {activeTab === 'payment_config' && (
            <div className="space-y-5">
              <h3 className="text-lg font-semibold text-gray-900">Payment Configuration</h3>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1.5">Default Due Day of Month</label>
                  <input type="number" min="1" max="28" value={defaultDueDay} onChange={e => setDefaultDueDay(e.target.value)}
                    className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500"
                    placeholder="1" />
                  <p className="text-xs text-gray-400 mt-1">e.g. 1 = 1st of each month</p>
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1.5">Late Fee Amount ($)</label>
                  <input type="number" step="0.01" min="0" value={lateFee} onChange={e => setLateFee(e.target.value)}
                    className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500"
                    placeholder="25.00" />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1.5">Grace Period (days)</label>
                  <input type="number" min="0" value={lateGraceDays} onChange={e => setLateGraceDays(e.target.value)}
                    className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500"
                    placeholder="5" />
                </div>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Payment Methods Enabled</label>
                <div className="flex gap-4">
                  {['card', 'ach'].map(method => (
                    <label key={method} className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={paymentMethods.includes(method)}
                        onChange={e => {
                          if (e.target.checked) setPaymentMethods([...paymentMethods, method])
                          else setPaymentMethods(paymentMethods.filter(m => m !== method))
                        }}
                        className="w-4 h-4 text-blue-600 rounded"
                      />
                      <span className="text-sm text-gray-700 uppercase font-medium">{method}</span>
                    </label>
                  ))}
                </div>
                <p className="text-xs text-gray-400 mt-1">Note: ACH requires Stripe account verification</p>
              </div>
              <button onClick={savePaymentConfig} disabled={saving}
                className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-400 text-white text-sm font-semibold rounded-xl transition-colors">
                {saving ? 'Saving...' : 'Save Payment Config'}
              </button>
            </div>
          )}

          {/* Documents */}
          {activeTab === 'documents_upload' && (
            <div className="space-y-5">
              <h3 className="text-lg font-semibold text-gray-900">Initial Document Setup</h3>
              <div className="p-4 bg-blue-50 rounded-xl border border-blue-100">
                <p className="text-sm text-blue-800">
                  Upload your HOA&apos;s foundational documents (bylaws, CC&Rs, rules & regulations) in the{' '}
                  <strong>Documents</strong> section. Once you&apos;ve uploaded your initial documents, mark this step complete.
                </p>
              </div>
              <a href="/kevin/admin/documents"
                className="inline-flex items-center gap-2 px-5 py-2.5 bg-gray-100 hover:bg-gray-200 text-gray-700 text-sm font-semibold rounded-xl transition-colors">
                <FileText size={16} />
                Go to Documents →
              </a>
              <div>
                <button onClick={markDocumentsDone}
                  className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold rounded-xl transition-colors">
                  Mark Documents as Uploaded
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
