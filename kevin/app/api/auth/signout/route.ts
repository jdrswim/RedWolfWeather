import { createServerSupabaseClient } from '@/lib/supabaseServer'
import { NextResponse } from 'next/server'

export async function POST() {
  const supabase = await createServerSupabaseClient()
  await supabase.auth.signOut()
  return NextResponse.redirect(new URL('/kevin/login', process.env.NEXT_PUBLIC_BASE_URL || 'http://localhost:3000'))
}
