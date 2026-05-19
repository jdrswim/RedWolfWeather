import { Building2 } from 'lucide-react'
import Link from 'next/link'

export default function AuthLayout({ children }: { children: React.ReactNode }) {
  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 via-white to-indigo-50 flex flex-col">
      <nav className="px-6 py-4">
        <Link href="/" className="inline-flex items-center gap-2 text-gray-900 hover:text-blue-600 transition-colors">
          <Building2 className="h-6 w-6 text-blue-600" />
          <span className="font-bold text-lg">HOA Manager</span>
        </Link>
      </nav>
      <div className="flex-1 flex items-center justify-center px-4 py-12">
        {children}
      </div>
    </div>
  )
}
