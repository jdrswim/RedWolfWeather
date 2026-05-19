'use client'

import { useState } from 'react'
import { CreditCard } from 'lucide-react'

interface PayDuesButtonProps {
  duesId: string
  amount: number
}

export default function PayDuesButton({ duesId, amount }: PayDuesButtonProps) {
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')

  async function handlePay() {
    setLoading(true)
    setError('')

    const res = await fetch('/kevin/api/stripe/checkout', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ duesId, amount }),
    })

    const data = await res.json()

    if (!res.ok || !data.url) {
      setError(data.error || 'Payment failed to initialize')
      setLoading(false)
      return
    }

    window.location.href = data.url
  }

  return (
    <div>
      {error && <p className="text-xs text-red-500 mb-1 text-right">{error}</p>}
      <button
        onClick={handlePay}
        disabled={loading}
        className="flex items-center gap-1.5 px-4 py-2 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-400 text-white text-sm font-semibold rounded-xl transition-colors whitespace-nowrap"
      >
        <CreditCard size={14} />
        {loading ? 'Processing...' : `Pay $${amount.toFixed(2)}`}
      </button>
    </div>
  )
}
