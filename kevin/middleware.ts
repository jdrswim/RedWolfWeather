import { createServerClient, type CookieOptions } from '@supabase/ssr'
import { NextResponse, type NextRequest } from 'next/server'

export async function middleware(request: NextRequest) {
  let supabaseResponse = NextResponse.next({ request })

  const supabase = createServerClient(
    process.env.NEXT_PUBLIC_SUPABASE_URL!,
    process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!,
    {
      cookies: {
        getAll() {
          return request.cookies.getAll()
        },
        setAll(cookiesToSet: { name: string; value: string; options: CookieOptions }[]) {
          cookiesToSet.forEach(({ name, value }) =>
            request.cookies.set(name, value)
          )
          supabaseResponse = NextResponse.next({ request })
          cookiesToSet.forEach(({ name, value, options }) =>
            supabaseResponse.cookies.set(name, value, options)
          )
        },
      },
    }
  )

  const { data: { user } } = await supabase.auth.getUser()
  const { pathname } = request.nextUrl

  // Protect dashboard routes
  const isDashboardRoute = pathname.startsWith('/kevin/owner') || pathname.startsWith('/kevin/admin')
  const isAuthRoute = pathname.startsWith('/kevin/login') || pathname.startsWith('/kevin/signup')
  const isSetupRoute = pathname.startsWith('/kevin/setup')

  if (isDashboardRoute && !user) {
    const loginUrl = new URL('/kevin/login', request.url)
    return NextResponse.redirect(loginUrl)
  }

  // If logged in and on auth route, redirect to appropriate dashboard
  if (isAuthRoute && user) {
    // Get user profile to determine role
    const { data: profile } = await supabase
      .from('profiles')
      .select('role')
      .eq('id', user.id)
      .single()

    const role = profile?.role || 'owner'
    const dashUrl = new URL(role === 'admin' ? '/kevin/admin' : '/kevin/owner', request.url)
    return NextResponse.redirect(dashUrl)
  }

  // Block owners from admin routes
  if (pathname.startsWith('/kevin/admin') && user) {
    const { data: profile } = await supabase
      .from('profiles')
      .select('role')
      .eq('id', user.id)
      .single()

    if (profile?.role !== 'admin') {
      return NextResponse.redirect(new URL('/kevin/owner', request.url))
    }
  }

  return supabaseResponse
}

export const config = {
  matcher: [
    '/kevin/((?!_next/static|_next/image|favicon.ico|.*\\.(?:svg|png|jpg|jpeg|gif|webp)$).*)',
  ],
}
