import { NextRequest, NextResponse } from 'next/server'
import { createServerSupabaseClient, createServiceRoleClient } from '@/lib/supabaseServer'

export async function POST(req: NextRequest) {
  const supabase = await createServerSupabaseClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 })

  const { data: profile } = await supabase.from('profiles').select('role').eq('id', user.id).single()
  if (profile?.role !== 'admin') return NextResponse.json({ error: 'Forbidden' }, { status: 403 })

  const { name, email, unitNumber } = await req.json()
  if (!name || !email) return NextResponse.json({ error: 'Name and email required' }, { status: 400 })

  const serviceClient = createServiceRoleClient()
  const { data: authData, error: authError } = await serviceClient.auth.admin.createUser({
    email,
    password: Math.random().toString(36).slice(-10) + '!Aa1',
    email_confirm: true,
    user_metadata: { name, role: 'owner' },
  })

  if (authError) return NextResponse.json({ error: authError.message }, { status: 400 })

  if (unitNumber) {
    await serviceClient.from('profiles').update({ unit_number: unitNumber }).eq('id', authData.user.id)
  }

  return NextResponse.json({ success: true, userId: authData.user.id })
}
