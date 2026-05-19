import { cn } from '@/lib/utils'

type BadgeVariant = 'success' | 'error' | 'warning' | 'info' | 'neutral'

interface BadgeProps {
  variant?: BadgeVariant
  children: React.ReactNode
  className?: string
}

const variants: Record<BadgeVariant, string> = {
  success: 'bg-green-50 text-green-700 border-green-100',
  error: 'bg-red-50 text-red-700 border-red-100',
  warning: 'bg-yellow-50 text-yellow-700 border-yellow-100',
  info: 'bg-blue-50 text-blue-700 border-blue-100',
  neutral: 'bg-gray-50 text-gray-600 border-gray-100',
}

export default function Badge({ variant = 'neutral', children, className }: BadgeProps) {
  return (
    <span
      className={cn(
        'inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium border',
        variants[variant],
        className
      )}
    >
      {children}
    </span>
  )
}

export function statusBadge(status: string) {
  const map: Record<string, BadgeVariant> = {
    paid: 'success',
    completed: 'success',
    unpaid: 'error',
    overdue: 'error',
    failed: 'error',
    partial: 'warning',
    pending: 'warning',
    refunded: 'info',
    read: 'neutral',
    unread: 'info',
  }
  return map[status] ?? 'neutral'
}
