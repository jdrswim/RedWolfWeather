import { LucideIcon } from 'lucide-react'
import clsx from 'clsx'

interface StatCardProps {
  label: string
  value: string | number
  icon: LucideIcon
  trend?: string
  trendUp?: boolean
  color?: 'blue' | 'green' | 'yellow' | 'red' | 'purple'
}

const colorMap = {
  blue:   { bg: 'bg-blue-50',   icon: 'bg-blue-600',   text: 'text-blue-600' },
  green:  { bg: 'bg-green-50',  icon: 'bg-green-600',  text: 'text-green-600' },
  yellow: { bg: 'bg-yellow-50', icon: 'bg-yellow-500', text: 'text-yellow-600' },
  red:    { bg: 'bg-red-50',    icon: 'bg-red-600',    text: 'text-red-600' },
  purple: { bg: 'bg-purple-50', icon: 'bg-purple-600', text: 'text-purple-600' },
}

export default function StatCard({ label, value, icon: Icon, trend, trendUp, color = 'blue' }: StatCardProps) {
  const c = colorMap[color]
  return (
    <div className={clsx('rounded-2xl p-6 shadow-sm border border-gray-100', c.bg)}>
      <div className="flex items-start justify-between">
        <div>
          <p className="text-sm font-medium text-gray-500 mb-1">{label}</p>
          <p className="text-3xl font-bold text-gray-900">{value}</p>
          {trend && (
            <p className={clsx('text-xs mt-1 font-medium', trendUp ? 'text-green-600' : 'text-red-500')}>
              {trend}
            </p>
          )}
        </div>
        <div className={clsx('p-3 rounded-xl', c.icon)}>
          <Icon className="w-6 h-6 text-white" />
        </div>
      </div>
    </div>
  )
}
