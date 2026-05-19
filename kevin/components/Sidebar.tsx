'use client'

import Link from 'next/link'
import { usePathname, useRouter } from 'next/navigation'
import { createClient } from '@/lib/supabaseClient'
import {
  LayoutDashboard, DollarSign, CreditCard, MessageSquare,
  FileText, Users, ReceiptIcon, LogOut, Home, Menu, X, Settings
} from 'lucide-react'
import { useState } from 'react'
import clsx from 'clsx'

interface SidebarProps {
  role: 'admin' | 'owner'
  userName: string
  unitNumber?: string | null
}

const adminLinks = [
  { href: '/kevin/admin', label: 'Dashboard', icon: LayoutDashboard, exact: true },
  { href: '/kevin/admin/owners', label: 'Owners', icon: Users },
  { href: '/kevin/admin/dues', label: 'Dues', icon: DollarSign },
  { href: '/kevin/admin/payments', label: 'Payments', icon: CreditCard },
  { href: '/kevin/admin/expenses', label: 'Expenses', icon: ReceiptIcon },
  { href: '/kevin/admin/messages', label: 'Messages', icon: MessageSquare },
  { href: '/kevin/admin/documents', label: 'Documents', icon: FileText },
  { href: '/kevin/admin/settings', label: 'Settings', icon: Settings },
]

const ownerLinks = [
  { href: '/kevin/owner', label: 'Dashboard', icon: LayoutDashboard, exact: true },
  { href: '/kevin/owner/dues', label: 'My Dues', icon: DollarSign },
  { href: '/kevin/owner/payments', label: 'Payments', icon: CreditCard },
  { href: '/kevin/owner/messages', label: 'Messages', icon: MessageSquare },
  { href: '/kevin/owner/documents', label: 'Documents', icon: FileText },
]

export default function Sidebar({ role, userName, unitNumber }: SidebarProps) {
  const pathname = usePathname()
  const router = useRouter()
  const [mobileOpen, setMobileOpen] = useState(false)
  const supabase = createClient()
  const links = role === 'admin' ? adminLinks : ownerLinks

  async function handleSignOut() {
    await supabase.auth.signOut()
    router.push('/kevin/login')
    router.refresh()
  }

  function isActive(href: string, exact?: boolean) {
    if (exact) return pathname === href
    return pathname.startsWith(href)
  }

  const SidebarContent = () => (
    <div className="flex flex-col h-full">
      {/* Logo */}
      <div className="px-6 py-5 border-b border-blue-700">
        <Link href={role === 'admin' ? '/kevin/admin' : '/kevin/owner'} className="flex items-center gap-3">
          <div className="w-9 h-9 bg-white/20 rounded-xl flex items-center justify-center">
            <Home className="w-5 h-5 text-white" />
          </div>
          <div>
            <div className="font-bold text-white text-sm leading-tight">Kevin&apos;s HOA</div>
            <div className="text-blue-200 text-xs capitalize">{role} Portal</div>
          </div>
        </Link>
      </div>

      {/* User info */}
      <div className="px-6 py-4 border-b border-blue-700">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-400 rounded-full flex items-center justify-center text-white font-bold text-sm">
            {userName.charAt(0).toUpperCase()}
          </div>
          <div className="min-w-0">
            <div className="text-white font-medium text-sm truncate">{userName}</div>
            {unitNumber && <div className="text-blue-200 text-xs">Unit {unitNumber}</div>}
          </div>
        </div>
      </div>

      {/* Nav links */}
      <nav className="flex-1 px-3 py-4 space-y-1 overflow-y-auto">
        {links.map(({ href, label, icon: Icon, exact }) => (
          <Link
            key={href}
            href={href}
            onClick={() => setMobileOpen(false)}
            className={clsx(
              'flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-medium transition-colors',
              isActive(href, exact)
                ? 'bg-white/20 text-white'
                : 'text-blue-100 hover:bg-white/10 hover:text-white'
            )}
          >
            <Icon className="w-4.5 h-4.5 flex-shrink-0" size={18} />
            {label}
          </Link>
        ))}
      </nav>

      {/* Sign out */}
      <div className="px-3 py-4 border-t border-blue-700">
        <button
          onClick={handleSignOut}
          className="w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-medium text-blue-100 hover:bg-white/10 hover:text-white transition-colors"
        >
          <LogOut size={18} />
          Sign out
        </button>
      </div>
    </div>
  )

  return (
    <>
      {/* Mobile toggle */}
      <button
        onClick={() => setMobileOpen(!mobileOpen)}
        className="lg:hidden fixed top-4 left-4 z-50 p-2 bg-blue-800 rounded-xl text-white shadow-lg"
      >
        {mobileOpen ? <X size={20} /> : <Menu size={20} />}
      </button>

      {/* Mobile overlay */}
      {mobileOpen && (
        <div
          className="lg:hidden fixed inset-0 bg-black/50 z-40"
          onClick={() => setMobileOpen(false)}
        />
      )}

      {/* Mobile sidebar */}
      <div className={clsx(
        'lg:hidden fixed left-0 top-0 h-full w-64 bg-gradient-to-b from-blue-900 to-blue-800 z-40 transition-transform duration-300',
        mobileOpen ? 'translate-x-0' : '-translate-x-full'
      )}>
        <SidebarContent />
      </div>

      {/* Desktop sidebar */}
      <div className="hidden lg:flex w-64 flex-shrink-0 flex-col bg-gradient-to-b from-blue-900 to-blue-800 min-h-screen">
        <SidebarContent />
      </div>
    </>
  )
}
