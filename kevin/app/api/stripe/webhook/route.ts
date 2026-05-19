import { NextRequest, NextResponse } from 'next/server'
import { stripe } from '@/lib/stripe'
import { createServiceRoleClient } from '@/lib/supabaseServer'
import Stripe from 'stripe'

export const config = {
  api: {
    bodyParser: false,
  },
}

export async function POST(req: NextRequest) {
  const body = await req.text()
  const sig = req.headers.get('stripe-signature')

  if (!sig || !process.env.STRIPE_WEBHOOK_SECRET) {
    return NextResponse.json({ error: 'Missing signature or secret' }, { status: 400 })
  }

  let event: Stripe.Event

  try {
    event = stripe.webhooks.constructEvent(body, sig, process.env.STRIPE_WEBHOOK_SECRET)
  } catch (err: any) {
    console.error('Webhook signature verification failed:', err.message)
    return NextResponse.json({ error: `Webhook Error: ${err.message}` }, { status: 400 })
  }

  const supabase = createServiceRoleClient()

  if (event.type === 'checkout.session.completed') {
    const session = event.data.object as Stripe.Checkout.Session
    const duesId = session.metadata?.dues_id
    const ownerId = session.metadata?.owner_id

    if (duesId && ownerId) {
      // Update payment status
      await supabase
        .from('payments')
        .update({
          status: 'completed',
          stripe_payment_intent_id: session.payment_intent as string,
        })
        .eq('stripe_session_id', session.id)

      // Get the payment amount
      const amountPaid = (session.amount_total || 0) / 100

      // Get current dues balance
      const { data: due } = await supabase
        .from('dues')
        .select('balance_remaining, amount_due')
        .eq('id', duesId)
        .single()

      if (due) {
        const newBalance = Math.max(0, due.balance_remaining - amountPaid)
        const newStatus = newBalance === 0 ? 'paid' : 'partial'

        await supabase
          .from('dues')
          .update({
            balance_remaining: newBalance,
            status: newStatus,
          })
          .eq('id', duesId)
      }
    }
  }

  if (event.type === 'payment_intent.payment_failed') {
    const paymentIntent = event.data.object as Stripe.PaymentIntent
    await supabase
      .from('payments')
      .update({ status: 'failed' })
      .eq('stripe_payment_intent_id', paymentIntent.id)
  }

  return NextResponse.json({ received: true })
}
