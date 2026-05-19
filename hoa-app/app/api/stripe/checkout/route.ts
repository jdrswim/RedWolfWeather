import { NextRequest, NextResponse } from 'next/server'
import { stripe } from '@/lib/stripe'
import { createClient } from '@/lib/supabaseServer'

export async function POST(req: NextRequest) {
  try {
    const supabase = await createClient()
    const { data: { user } } = await supabase.auth.getUser()

    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const { dueId, amount } = await req.json()

    if (!dueId || !amount || amount <= 0) {
      return NextResponse.json({ error: 'Invalid request' }, { status: 400 })
    }

    const { data: due } = await supabase
      .from('dues')
      .select('*, profiles(*)')
      .eq('id', dueId)
      .eq('owner_id', user.id)
      .single()

    if (!due) {
      return NextResponse.json({ error: 'Due not found' }, { status: 404 })
    }

    const { data: settings } = await supabase.from('hoa_settings').select('hoa_name, payment_methods').single()

    const paymentMethodTypes: ('card' | 'us_bank_account')[] = ['card']
    if (settings?.payment_methods?.includes('ach')) {
      paymentMethodTypes.push('us_bank_account')
    }

    const appUrl = process.env.NEXT_PUBLIC_APP_URL || 'http://localhost:3000'

    const session = await stripe.checkout.sessions.create({
      payment_method_types: paymentMethodTypes,
      line_items: [
        {
          price_data: {
            currency: 'usd',
            product_data: {
              name: `HOA Dues — ${due.month_year}`,
              description: `${settings?.hoa_name ?? 'HOA'} · Unit ${due.profiles?.unit_number ?? 'N/A'}`,
            },
            unit_amount: Math.round(amount * 100),
          },
          quantity: 1,
        },
      ],
      mode: 'payment',
      success_url: `${appUrl}/dashboard/owner/dues?success=1&session_id={CHECKOUT_SESSION_ID}`,
      cancel_url: `${appUrl}/dashboard/owner/dues?cancelled=1`,
      metadata: {
        due_id: dueId,
        owner_id: user.id,
        amount: String(amount),
      },
      customer_email: due.profiles?.email ?? undefined,
    })

    // Pre-create a pending payment record
    await supabase.from('payments').insert({
      owner_id: user.id,
      amount,
      stripe_session_id: session.id,
      status: 'pending',
      payment_date: new Date().toISOString(),
      due_id: dueId,
    })

    return NextResponse.json({ url: session.url })
  } catch (error: any) {
    console.error('Stripe checkout error:', error)
    return NextResponse.json({ error: error.message ?? 'Stripe error' }, { status: 500 })
  }
}
