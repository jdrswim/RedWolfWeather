import { NextRequest, NextResponse } from 'next/server'
import { stripe } from '@/lib/stripe'
import { createServiceClient } from '@/lib/supabaseServer'
import Stripe from 'stripe'

export async function POST(req: NextRequest) {
  const body = await req.text()
  const signature = req.headers.get('stripe-signature')

  if (!signature) {
    return NextResponse.json({ error: 'No signature' }, { status: 400 })
  }

  let event: Stripe.Event

  try {
    event = stripe.webhooks.constructEvent(body, signature, process.env.STRIPE_WEBHOOK_SECRET!)
  } catch (err: any) {
    console.error('Webhook signature verification failed:', err.message)
    return NextResponse.json({ error: `Webhook error: ${err.message}` }, { status: 400 })
  }

  const supabase = await createServiceClient()

  switch (event.type) {
    case 'checkout.session.completed': {
      const session = event.data.object as Stripe.Checkout.Session
      const { due_id, owner_id, amount } = session.metadata ?? {}

      if (!due_id || !owner_id) break

      // Update payment record to completed
      await supabase
        .from('payments')
        .update({
          status: 'completed',
          stripe_payment_intent_id: session.payment_intent as string,
          payment_date: new Date().toISOString(),
        })
        .eq('stripe_session_id', session.id)

      // Update the due status
      const paidAmount = parseFloat(amount ?? '0')
      const { data: due } = await supabase.from('dues').select('amount_due, balance_remaining').eq('id', due_id).single()

      if (due) {
        const newBalance = Math.max(0, due.balance_remaining - paidAmount)
        await supabase
          .from('dues')
          .update({
            balance_remaining: newBalance,
            status: newBalance <= 0 ? 'paid' : 'partial',
          })
          .eq('id', due_id)
      }
      break
    }

    case 'checkout.session.expired':
    case 'payment_intent.payment_failed': {
      const session = event.data.object as Stripe.Checkout.Session
      if (session.id) {
        await supabase
          .from('payments')
          .update({ status: 'failed' })
          .eq('stripe_session_id', session.id)
      }
      break
    }

    default:
      // Unhandled event types are fine to ignore
      break
  }

  return NextResponse.json({ received: true })
}
