'use client'

import { useEffect, useState } from 'react'
import { createClient } from '@/lib/supabaseClient'
import AdminSidebar from '@/components/layout/AdminSidebar'
import PageHeader from '@/components/layout/PageHeader'
import Button from '@/components/ui/Button'
import { AlertCircle, CheckCircle2, ExternalLink } from 'lucide-react'

export default function SettingsPage() {
  const [settings, setSettings] = useState({
    hoaName: '', address: '', city: '', state: '', zip: '',
    numUnits: '', defaultDues: '', defaultDueDay: '1',
    lateFeeAmount: '', lateFeeGraceDays: '5',
    enableCard: true, enableACH: false,
    accountingStartDate: '',
  })
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState(false)
  const [hoaName, setHoaName] = useState('')
  const supabase = createClient()

  useEffect(() => {
    supabase.from('hoa_settings').select('*').single().then(({ data }) => {
      if (data) {
        setSettings({
          hoaName: data.hoa_name ?? '',
          address: data.address ?? '',
          city: data.city ?? '',
          state: data.state ?? '',
          zip: data.zip ?? '',
          numUnits: String(data.num_units ?? ''),
          defaultDues: String(data.default_monthly_dues ?? ''),
          defaultDueDay: String(data.default_due_day ?? '1'),
          lateFeeAmount: String(data.late_fee_amount ?? ''),
          lateFeeGraceDays: String(data.late_fee_days_grace ?? '5'),
          enableCard: data.payment_methods?.includes('card') ?? true,
          enableACH: data.payment_methods?.includes('ach') ?? false,
          accountingStartDate: data.accounting_start_date ?? '',
        })
        setHoaName(data.hoa_name ?? '')
      }
      setLoading(false)
    })
  }, [supabase])

  async function handleSave(e: React.FormEvent) {
    e.preventDefault()
    setSaving(true)
    await supabase.from('hoa_settings').upsert({
      id: '00000000-0000-0000-0000-000000000001',
      hoa_name: settings.hoaName,
      address: settings.address,
      city: settings.city,
      state: settings.state,
      zip: settings.zip,
      num_units: parseInt(settings.numUnits) || 0,
      default_monthly_dues: settings.defaultDues ? parseFloat(settings.defaultDues) : null,
      default_due_day: parseInt(settings.defaultDueDay) || 1,
      late_fee_amount: settings.lateFeeAmount ? parseFloat(settings.lateFeeAmount) : null,
      late_fee_days_grace: parseInt(settings.lateFeeGraceDays) || 5,
      payment_methods: [
        ...(settings.enableCard ? ['card'] : []),
        ...(settings.enableACH ? ['ach'] : []),
      ],
      accounting_start_date: settings.accountingStartDate || null,
      updated_at: new Date().toISOString(),
    })
    setHoaName(settings.hoaName)
    setSaving(false)
    setSaved(true)
    setTimeout(() => setSaved(false), 3000)
  }

  function set(field: string, value: string | boolean) {
    setSettings((prev) => ({ ...prev, [field]: value }))
  }

  if (loading) {
    return (
      <div className="flex min-h-screen bg-gray-50">
        <AdminSidebar hoaName={hoaName} />
        <main className="flex-1 ml-64 p-8">
          <div className="text-center text-gray-400 py-20">Loading settings…</div>
        </main>
      </div>
    )
  }

  return (
    <div className="flex min-h-screen bg-gray-50">
      <AdminSidebar hoaName={hoaName} />
      <main className="flex-1 ml-64 p-8">
        <div className="max-w-3xl mx-auto">
          <PageHeader title="Settings" description="Configure your HOA organization and payment rules" />

          <form onSubmit={handleSave} className="space-y-6">
            {/* Organization */}
            <Section title="Organization">
              <Field label="HOA name *" value={settings.hoaName} onChange={(v) => set('hoaName', v)} placeholder="Sunset Ridge HOA" />
              <Field label="Street address" value={settings.address} onChange={(v) => set('address', v)} placeholder="123 Main Street" />
              <div className="grid grid-cols-3 gap-3">
                <Field label="City" value={settings.city} onChange={(v) => set('city', v)} placeholder="Austin" />
                <Field label="State" value={settings.state} onChange={(v) => set('state', v)} placeholder="TX" />
                <Field label="ZIP" value={settings.zip} onChange={(v) => set('zip', v)} placeholder="78701" />
              </div>
              <div className="grid grid-cols-2 gap-3">
                <Field label="Number of units" value={settings.numUnits} onChange={(v) => set('numUnits', v)} type="number" placeholder="24" />
                <Field label="Default monthly dues ($)" value={settings.defaultDues} onChange={(v) => set('defaultDues', v)} type="number" placeholder="200.00" />
              </div>
            </Section>

            {/* Stripe */}
            <Section title="Stripe / Payment Setup">
              <div className="bg-blue-50 border border-blue-100 rounded-xl p-4 mb-4">
                <div className="flex gap-3">
                  <AlertCircle className="h-5 w-5 text-blue-600 flex-shrink-0 mt-0.5" />
                  <div>
                    <p className="text-sm font-medium text-blue-900">Stripe API keys are stored as environment variables</p>
                    <p className="text-xs text-blue-700 mt-1">
                      For security, Stripe keys are <strong>not</strong> stored in the database.
                      Set them in your <code className="bg-blue-100 px-1 rounded">.env.local</code> file
                      or your Vercel project environment variables.
                    </p>
                    <a
                      href="https://dashboard.stripe.com/apikeys"
                      target="_blank"
                      rel="noreferrer"
                      className="inline-flex items-center gap-1 text-xs text-blue-600 font-medium mt-2 hover:text-blue-700"
                    >
                      Open Stripe Dashboard
                      <ExternalLink className="h-3 w-3" />
                    </a>
                  </div>
                </div>
              </div>
              <div className="bg-gray-50 rounded-xl border border-gray-100 p-4 font-mono text-xs text-gray-700 space-y-1">
                <p>NEXT_PUBLIC_STRIPE_PUBLIC_KEY=pk_live_...</p>
                <p>STRIPE_SECRET_KEY=sk_live_...</p>
                <p>STRIPE_WEBHOOK_SECRET=whsec_...</p>
              </div>
              <div className="bg-yellow-50 border border-yellow-100 rounded-xl p-4 mt-3">
                <p className="text-xs text-yellow-800">
                  <strong>Bank payouts</strong> are configured in your Stripe Dashboard under
                  <strong> Settings → Bank accounts and scheduling</strong>.
                  This application does not automatically connect to or manage your bank account.
                </p>
              </div>
            </Section>

            {/* Payment rules */}
            <Section title="Payment Rules">
              <Field
                label="Dues due on day of month"
                value={settings.defaultDueDay}
                onChange={(v) => set('defaultDueDay', v)}
                type="number"
                hint="1–28. Dues are due on this day each month."
              />
              <div className="grid grid-cols-2 gap-3">
                <Field
                  label="Late fee ($)"
                  value={settings.lateFeeAmount}
                  onChange={(v) => set('lateFeeAmount', v)}
                  type="number"
                  placeholder="25.00"
                  hint="Leave blank to disable"
                />
                <Field
                  label="Grace period (days)"
                  value={settings.lateFeeGraceDays}
                  onChange={(v) => set('lateFeeGraceDays', v)}
                  type="number"
                  placeholder="5"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Accepted payment methods</label>
                <div className="space-y-2">
                  <label className="flex items-center gap-3 p-3 rounded-lg border border-gray-100 cursor-pointer hover:bg-gray-50">
                    <input type="checkbox" checked={settings.enableCard}
                      onChange={(e) => set('enableCard', e.target.checked)}
                      className="rounded border-gray-300 text-blue-600" />
                    <span className="text-sm font-medium text-gray-900">Credit / Debit card</span>
                  </label>
                  <label className="flex items-center gap-3 p-3 rounded-lg border border-gray-100 cursor-pointer hover:bg-gray-50">
                    <input type="checkbox" checked={settings.enableACH}
                      onChange={(e) => set('enableACH', e.target.checked)}
                      className="rounded border-gray-300 text-blue-600" />
                    <span className="text-sm font-medium text-gray-900">ACH bank transfer</span>
                  </label>
                </div>
              </div>
            </Section>

            {/* Financial */}
            <Section title="Financial">
              <Field
                label="Accounting start date"
                value={settings.accountingStartDate}
                onChange={(v) => set('accountingStartDate', v)}
                type="date"
              />
            </Section>

            <div className="flex items-center gap-4">
              <Button type="submit" size="lg" loading={saving}>
                Save settings
              </Button>
              {saved && (
                <div className="flex items-center gap-1.5 text-sm text-green-600">
                  <CheckCircle2 className="h-4 w-4" />
                  Saved
                </div>
              )}
            </div>
          </form>
        </div>
      </main>
    </div>
  )
}

function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div className="bg-white rounded-xl border border-gray-100 shadow-sm p-6">
      <h3 className="font-semibold text-gray-900 mb-5">{title}</h3>
      <div className="space-y-4">{children}</div>
    </div>
  )
}

function Field({ label, value, onChange, placeholder, type = 'text', hint }: {
  label: string; value: string; onChange: (v: string) => void
  placeholder?: string; type?: string; hint?: string
}) {
  return (
    <div>
      <label className="block text-sm font-medium text-gray-700 mb-1.5">{label}</label>
      <input type={type} value={value} onChange={(e) => onChange(e.target.value)} placeholder={placeholder}
        className="w-full px-3.5 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
      {hint && <p className="text-xs text-gray-400 mt-1">{hint}</p>}
    </div>
  )
}
