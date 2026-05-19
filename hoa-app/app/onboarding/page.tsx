'use client'

import { useState } from 'react'
import { useRouter } from 'next/navigation'
import { createClient } from '@/lib/supabaseClient'
import {
  Building2,
  CreditCard,
  Settings,
  User,
  FileText,
  ChevronRight,
  ChevronLeft,
  CheckCircle2,
  AlertCircle,
} from 'lucide-react'
import Button from '@/components/ui/Button'

type Step = 'organization' | 'financial' | 'payment' | 'identity' | 'review'

const steps: { id: Step; label: string; icon: React.ElementType }[] = [
  { id: 'organization', label: 'Organization', icon: Building2 },
  { id: 'financial', label: 'Financial', icon: CreditCard },
  { id: 'payment', label: 'Payment Rules', icon: Settings },
  { id: 'identity', label: 'Admin Profile', icon: User },
  { id: 'review', label: 'Review & Launch', icon: CheckCircle2 },
]

export default function OnboardingPage() {
  const router = useRouter()
  const [currentStep, setCurrentStep] = useState<Step>('organization')
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState('')

  const [org, setOrg] = useState({
    hoaName: '',
    address: '',
    city: '',
    state: '',
    zip: '',
    numUnits: '',
    defaultDues: '',
  })

  const [financial, setFinancial] = useState({
    accountingStartDate: '',
    stripePublicKey: '',
    stripeSecretKey: '',
    stripeWebhookSecret: '',
  })

  const [payment, setPayment] = useState({
    defaultDueDay: '1',
    lateFeeAmount: '',
    lateFeeGraceDays: '5',
    enableCard: true,
    enableACH: false,
  })

  const [adminProfile, setAdminProfile] = useState({
    adminName: '',
    adminEmail: '',
  })

  const stepIndex = steps.findIndex((s) => s.id === currentStep)

  function next() {
    const idx = steps.findIndex((s) => s.id === currentStep)
    if (idx < steps.length - 1) setCurrentStep(steps[idx + 1].id)
  }

  function back() {
    const idx = steps.findIndex((s) => s.id === currentStep)
    if (idx > 0) setCurrentStep(steps[idx - 1].id)
  }

  async function handleFinish() {
    setSaving(true)
    setError('')

    const supabase = createClient()
    const { data: { user } } = await supabase.auth.getUser()
    if (!user) { router.push('/login'); return }

    const { error: upsertError } = await supabase.from('hoa_settings').upsert({
      id: '00000000-0000-0000-0000-000000000001',
      hoa_name: org.hoaName,
      address: org.address,
      city: org.city,
      state: org.state,
      zip: org.zip,
      num_units: parseInt(org.numUnits) || 0,
      default_monthly_dues: org.defaultDues ? parseFloat(org.defaultDues) : null,
      accounting_start_date: financial.accountingStartDate || null,
      default_due_day: parseInt(payment.defaultDueDay) || 1,
      late_fee_amount: payment.lateFeeAmount ? parseFloat(payment.lateFeeAmount) : null,
      late_fee_days_grace: parseInt(payment.lateFeeGraceDays) || 5,
      payment_methods: [
        ...(payment.enableCard ? ['card'] : []),
        ...(payment.enableACH ? ['ach'] : []),
      ],
      stripe_configured: !!(financial.stripePublicKey && financial.stripeSecretKey),
      onboarding_completed: true,
    })

    if (upsertError) {
      setError(upsertError.message)
      setSaving(false)
      return
    }

    // Update admin profile name if provided
    if (adminProfile.adminName) {
      await supabase
        .from('profiles')
        .update({ name: adminProfile.adminName })
        .eq('id', user.id)
    }

    router.push('/dashboard/admin')
    router.refresh()
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 via-white to-indigo-50">
      <div className="max-w-2xl mx-auto px-4 py-12">
        {/* Header */}
        <div className="text-center mb-10">
          <div className="inline-flex items-center gap-2 mb-4">
            <Building2 className="h-8 w-8 text-blue-600" />
            <span className="text-2xl font-bold text-gray-900">HOA Manager</span>
          </div>
          <h1 className="text-3xl font-bold text-gray-900">Let&apos;s set up your HOA</h1>
          <p className="text-gray-500 mt-2">Complete these steps to get started. You can update everything later.</p>
        </div>

        {/* Step indicators */}
        <div className="flex items-center gap-2 mb-8">
          {steps.map((step, idx) => {
            const active = step.id === currentStep
            const done = idx < stepIndex
            return (
              <div key={step.id} className="flex items-center flex-1">
                <button
                  onClick={() => done && setCurrentStep(step.id)}
                  className={`flex items-center gap-2 flex-1 ${done ? 'cursor-pointer' : 'cursor-default'}`}
                >
                  <div
                    className={`h-8 w-8 rounded-full flex items-center justify-center flex-shrink-0 text-xs font-bold transition-all ${
                      done
                        ? 'bg-green-500 text-white'
                        : active
                        ? 'bg-blue-600 text-white'
                        : 'bg-gray-100 text-gray-400'
                    }`}
                  >
                    {done ? <CheckCircle2 className="h-4 w-4" /> : idx + 1}
                  </div>
                  {idx < steps.length - 1 && (
                    <div className={`flex-1 h-0.5 mx-1 ${done ? 'bg-green-400' : 'bg-gray-100'}`} />
                  )}
                </button>
              </div>
            )
          })}
        </div>

        {/* Step content */}
        <div className="bg-white rounded-2xl border border-gray-100 shadow-xl shadow-gray-100/50 p-8">
          <div className="flex items-center gap-3 mb-6">
            {(() => {
              const s = steps[stepIndex]
              return (
                <>
                  <div className="h-10 w-10 bg-blue-50 rounded-xl flex items-center justify-center">
                    <s.icon className="h-5 w-5 text-blue-600" />
                  </div>
                  <div>
                    <p className="text-xs text-gray-400 font-medium">Step {stepIndex + 1} of {steps.length}</p>
                    <h2 className="text-lg font-semibold text-gray-900">{s.label}</h2>
                  </div>
                </>
              )
            })()}
          </div>

          {currentStep === 'organization' && (
            <div className="space-y-4">
              <Field label="HOA / Community name *" value={org.hoaName} onChange={(v) => setOrg({ ...org, hoaName: v })} placeholder="Sunset Ridge HOA" />
              <Field label="Street address *" value={org.address} onChange={(v) => setOrg({ ...org, address: v })} placeholder="123 Main Street" />
              <div className="grid grid-cols-3 gap-3">
                <Field label="City *" value={org.city} onChange={(v) => setOrg({ ...org, city: v })} placeholder="Austin" />
                <Field label="State *" value={org.state} onChange={(v) => setOrg({ ...org, state: v })} placeholder="TX" />
                <Field label="ZIP *" value={org.zip} onChange={(v) => setOrg({ ...org, zip: v })} placeholder="78701" />
              </div>
              <div className="grid grid-cols-2 gap-3">
                <Field label="Number of units *" value={org.numUnits} onChange={(v) => setOrg({ ...org, numUnits: v })} placeholder="24" type="number" />
                <Field label="Default monthly dues ($)" value={org.defaultDues} onChange={(v) => setOrg({ ...org, defaultDues: v })} placeholder="200.00" type="number" />
              </div>
            </div>
          )}

          {currentStep === 'financial' && (
            <div className="space-y-5">
              <div className="bg-blue-50 border border-blue-100 rounded-xl p-4">
                <div className="flex gap-3">
                  <AlertCircle className="h-5 w-5 text-blue-600 flex-shrink-0 mt-0.5" />
                  <div>
                    <p className="text-sm font-medium text-blue-900">About Stripe API keys</p>
                    <p className="text-xs text-blue-700 mt-1">
                      Get your keys from{' '}
                      <strong>dashboard.stripe.com → Developers → API keys</strong>.
                      For security, these should be stored as environment variables
                      (<code className="bg-blue-100 px-1 rounded">STRIPE_SECRET_KEY</code>,{' '}
                      <code className="bg-blue-100 px-1 rounded">NEXT_PUBLIC_STRIPE_PUBLIC_KEY</code>).
                      Enter them here to verify your setup — they are not stored in the database.
                    </p>
                  </div>
                </div>
              </div>
              <Field
                label="Stripe Publishable Key (pk_...)"
                value={financial.stripePublicKey}
                onChange={(v) => setFinancial({ ...financial, stripePublicKey: v })}
                placeholder="pk_test_..."
              />
              <Field
                label="Stripe Secret Key (sk_...)"
                value={financial.stripeSecretKey}
                onChange={(v) => setFinancial({ ...financial, stripeSecretKey: v })}
                placeholder="sk_test_..."
                type="password"
              />
              <Field
                label="Stripe Webhook Secret (whsec_...)"
                value={financial.stripeWebhookSecret}
                onChange={(v) => setFinancial({ ...financial, stripeWebhookSecret: v })}
                placeholder="whsec_..."
                type="password"
              />
              <div className="bg-yellow-50 border border-yellow-100 rounded-xl p-4">
                <p className="text-xs text-yellow-800">
                  <strong>Bank payouts:</strong> Payouts are handled directly by Stripe to your connected bank account.
                  Set up your bank payout destination in your Stripe Dashboard under
                  <strong> Settings → Bank accounts and scheduling</strong>.
                  This app does not automatically connect to your bank — that must be configured in Stripe.
                </p>
              </div>
              <Field
                label="Accounting start date"
                value={financial.accountingStartDate}
                onChange={(v) => setFinancial({ ...financial, accountingStartDate: v })}
                type="date"
              />
            </div>
          )}

          {currentStep === 'payment' && (
            <div className="space-y-5">
              <Field
                label="Monthly dues due on day of month"
                value={payment.defaultDueDay}
                onChange={(v) => setPayment({ ...payment, defaultDueDay: v })}
                placeholder="1"
                type="number"
                hint="Enter 1–28. Dues will be due on this day each month."
              />
              <div className="grid grid-cols-2 gap-3">
                <Field
                  label="Late fee amount ($)"
                  value={payment.lateFeeAmount}
                  onChange={(v) => setPayment({ ...payment, lateFeeAmount: v })}
                  placeholder="25.00"
                  type="number"
                  hint="Leave blank to disable"
                />
                <Field
                  label="Grace period (days)"
                  value={payment.lateFeeGraceDays}
                  onChange={(v) => setPayment({ ...payment, lateFeeGraceDays: v })}
                  placeholder="5"
                  type="number"
                />
              </div>
              <div>
                <p className="text-sm font-medium text-gray-700 mb-3">Payment methods</p>
                <div className="space-y-2">
                  <label className="flex items-center gap-3 p-3 rounded-lg border border-gray-100 cursor-pointer hover:bg-gray-50 transition-colors">
                    <input
                      type="checkbox"
                      checked={payment.enableCard}
                      onChange={(e) => setPayment({ ...payment, enableCard: e.target.checked })}
                      className="rounded border-gray-300 text-blue-600"
                    />
                    <div>
                      <p className="text-sm font-medium text-gray-900">Credit / Debit card</p>
                      <p className="text-xs text-gray-400">2.9% + $0.30 per transaction (Stripe fees)</p>
                    </div>
                  </label>
                  <label className="flex items-center gap-3 p-3 rounded-lg border border-gray-100 cursor-pointer hover:bg-gray-50 transition-colors">
                    <input
                      type="checkbox"
                      checked={payment.enableACH}
                      onChange={(e) => setPayment({ ...payment, enableACH: e.target.checked })}
                      className="rounded border-gray-300 text-blue-600"
                    />
                    <div>
                      <p className="text-sm font-medium text-gray-900">ACH bank transfer</p>
                      <p className="text-xs text-gray-400">0.8% capped at $5 (Stripe fees) — requires Stripe ACH setup</p>
                    </div>
                  </label>
                </div>
              </div>
            </div>
          )}

          {currentStep === 'identity' && (
            <div className="space-y-4">
              <div className="bg-green-50 border border-green-100 rounded-xl p-4">
                <p className="text-sm text-green-800">
                  Your account was created during signup. Update your display name below if needed.
                </p>
              </div>
              <Field
                label="Your display name"
                value={adminProfile.adminName}
                onChange={(v) => setAdminProfile({ ...adminProfile, adminName: v })}
                placeholder="Jane Smith"
              />
            </div>
          )}

          {currentStep === 'review' && (
            <div className="space-y-4">
              <p className="text-sm text-gray-500">Review your setup before launching your HOA portal.</p>

              <div className="space-y-3">
                <ReviewRow label="HOA Name" value={org.hoaName || '—'} />
                <ReviewRow label="Address" value={[org.address, org.city, org.state, org.zip].filter(Boolean).join(', ') || '—'} />
                <ReviewRow label="Units" value={org.numUnits || '—'} />
                <ReviewRow label="Default dues" value={org.defaultDues ? `$${org.defaultDues}/month` : 'Not set'} />
                <ReviewRow label="Stripe configured" value={financial.stripePublicKey ? 'Yes' : 'No (add to .env later)'} />
                <ReviewRow label="Due day" value={`Day ${payment.defaultDueDay} of each month`} />
                <ReviewRow label="Late fee" value={payment.lateFeeAmount ? `$${payment.lateFeeAmount} after ${payment.lateFeeGraceDays} days` : 'None'} />
                <ReviewRow label="Payment methods" value={[payment.enableCard && 'Card', payment.enableACH && 'ACH'].filter(Boolean).join(', ') || 'None'} />
              </div>

              {error && (
                <div className="bg-red-50 border border-red-100 text-red-700 text-sm px-4 py-3 rounded-lg">
                  {error}
                </div>
              )}
            </div>
          )}

          {/* Navigation */}
          <div className="flex items-center justify-between mt-8 pt-6 border-t border-gray-100">
            <Button
              variant="outline"
              onClick={back}
              disabled={stepIndex === 0}
            >
              <ChevronLeft className="h-4 w-4" />
              Back
            </Button>

            {currentStep === 'review' ? (
              <Button onClick={handleFinish} loading={saving} size="lg">
                Launch my HOA portal
                <ChevronRight className="h-4 w-4" />
              </Button>
            ) : (
              <Button onClick={next} disabled={currentStep === 'organization' && !org.hoaName}>
                Continue
                <ChevronRight className="h-4 w-4" />
              </Button>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}

function Field({
  label,
  value,
  onChange,
  placeholder,
  type = 'text',
  hint,
}: {
  label: string
  value: string
  onChange: (v: string) => void
  placeholder?: string
  type?: string
  hint?: string
}) {
  return (
    <div>
      <label className="block text-sm font-medium text-gray-700 mb-1.5">{label}</label>
      <input
        type={type}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        className="w-full px-3.5 py-2.5 border border-gray-200 rounded-lg text-gray-900 placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
      />
      {hint && <p className="text-xs text-gray-400 mt-1">{hint}</p>}
    </div>
  )
}

function ReviewRow({ label, value }: { label: string; value: string }) {
  return (
    <div className="flex items-center justify-between py-2 border-b border-gray-50 last:border-0">
      <span className="text-sm text-gray-500">{label}</span>
      <span className="text-sm font-medium text-gray-900">{value}</span>
    </div>
  )
}
