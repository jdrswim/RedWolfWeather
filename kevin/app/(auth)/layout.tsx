import { Home } from 'lucide-react'
import Link from 'next/link'

export default function AuthLayout({ children }: { children: React.ReactNode }) {
  return (
    <div className="min-h-screen flex flex-col bg-gradient-to-br from-blue-900 via-blue-800 to-blue-700">
      <div className="flex items-center p-6">
        <Link href="/kevin" className="flex items-center gap-2 text-white/80 hover:text-white transition-colors">
          <Home className="w-5 h-5" />
          <span className="font-semibold text-lg">Kevin&apos;s HOA</span>
        </Link>
      </div>
      <div className="flex-1 flex items-center justify-center p-6">
        {children}
      </div>
      <div className="p-6 text-center text-white/50 text-sm">
        © {new Date().getFullYear()} Kevin&apos;s HOA Management. Powered by RedWolfWeather.com
      </div>
    </div>
  )
}
