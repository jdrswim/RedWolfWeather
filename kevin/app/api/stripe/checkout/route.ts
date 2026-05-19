import { NextRequest, NextResponse } from 'next/server'
import { stripe } from '@/lib/stripe'
import { createServerSupabaseClient } from '@/lib/supabaseServer'

export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient()
    const { data: { user } } = await supabase.auth.getUser()

    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })
    }

    const { duesId, amount } = await req.json()

    if (!duesId || !amount) {
      return NextResponse.json({ error: 'Missing duesId or amount' }, { status: 400 })
    }

    const { data: profile } = await supabase
      .from('profiles')
      .select('name, email')
      .eq('id', user.id)
      .single()

    const baseUrl = process.env.NEXT_PUBLIC_BASE_URL || 'http://localhost:3000'

    const session = await stripe.checkout.sessions.create({
      payment_method_types: ['card'],
      line_items: [
        {
          price_data: {
            currency: 'usd',
            product_data: {
              name: "HOA Dues Payment",
              description: `Payment for HOA dues - ${profile?.name || user.email}`,
            },
            unit_amount: Math.round(amount * 100),
          },
          quantity: 1,
        },
      ],
      mode: 'payment',
      success_url: `${baseUrl}/kevin/owner/payments?success=true&session_id={CHECKOUT_SESSION_ID}`,
      cancel_url: `${baseUrl}/kevin/owner/dues?cancelled=true`,
      customer_email: profile?.email || user.email || undefined,
      metadata: {
        dues_id: duesId,
        owner_id: user.id,
      },
    })

    // Create a pending payment record
    await supabase.from('payments').insert({
      owner_id: user.id,
      dues_id: duesId,
      amount,
      stripe_session_id: session.id,
      status: 'pending',
    })

    return NextResponse.json({ url: session.url })
  } catch (error: any) {
    console.error('Checkout error:', error)
    return NextResponse.json({ error: error.message }, { status: 500 })
  }
}
