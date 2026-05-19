'use client'

import { CheckCircle2, Circle, ChevronRight } from 'lucide-react'
import Link from 'next/link'
import { cn } from '@/lib/utils'

export interface ChecklistItem {
  id: string
  label: string
  description: string
  completed: boolean
  href: string
}

interface SetupChecklistProps {
  items: ChecklistItem[]
  className?: string
}

export default function SetupChecklist({ items, className }: SetupChecklistProps) {
  const completed = items.filter((i) => i.completed).length
  const total = items.length
  const pct = Math.round((completed / total) * 100)

  return (
    <div className={cn('bg-white rounded-xl border border-gray-100 shadow-sm', className)}>
      <div className="px-6 py-4 border-b border-gray-100">
        <div className="flex items-center justify-between mb-3">
          <h3 className="font-semibold text-gray-900">Setup Checklist</h3>
          <span className="text-sm font-medium text-gray-500">{completed}/{total}</span>
        </div>
        <div className="w-full bg-gray-100 rounded-full h-2">
          <div
            className="bg-blue-600 h-2 rounded-full transition-all duration-500"
            style={{ width: `${pct}%` }}
          />
        </div>
        <p className="text-xs text-gray-400 mt-2">{pct}% complete</p>
      </div>

      <ul className="divide-y divide-gray-50">
        {items.map((item) => (
          <li key={item.id}>
            <Link
              href={item.href}
              className={cn(
                'flex items-center gap-4 px-6 py-4 hover:bg-gray-50 transition-colors group',
                item.completed && 'opacity-60'
              )}
            >
              {item.completed ? (
                <CheckCircle2 className="h-5 w-5 text-green-500 flex-shrink-0" />
              ) : (
                <Circle className="h-5 w-5 text-gray-300 flex-shrink-0 group-hover:text-blue-400 transition-colors" />
              )}
              <div className="flex-1 min-w-0">
                <p className={cn('text-sm font-medium', item.completed ? 'text-gray-400 line-through' : 'text-gray-900')}>
                  {item.label}
                </p>
                <p className="text-xs text-gray-400 mt-0.5 truncate">{item.description}</p>
              </div>
              {!item.completed && (
                <ChevronRight className="h-4 w-4 text-gray-300 group-hover:text-blue-500 transition-colors flex-shrink-0" />
              )}
            </Link>
          </li>
        ))}
      </ul>

      {completed === total && (
        <div className="px-6 py-4 bg-green-50 border-t border-green-100 rounded-b-xl">
          <p className="text-sm font-medium text-green-700 text-center">
            All setup steps complete! Your HOA is ready to go.
          </p>
        </div>
      )}
    </div>
  )
}
