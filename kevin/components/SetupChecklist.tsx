'use client'

import { CheckCircle2, Circle, ChevronRight } from 'lucide-react'
import Link from 'next/link'

interface ChecklistItem {
  step_key: string
  step_label: string
  completed: boolean
  sort_order: number
}

interface SetupChecklistProps {
  items: ChecklistItem[]
}

export default function SetupChecklist({ items }: SetupChecklistProps) {
  const sorted = [...items].sort((a, b) => a.sort_order - b.sort_order)
  const completed = sorted.filter(i => i.completed).length
  const total = sorted.length
  const allDone = completed === total

  return (
    <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
      <div className="px-6 py-5 border-b border-gray-100">
        <div className="flex items-center justify-between mb-3">
          <h2 className="text-lg font-semibold text-gray-900">Setup Checklist</h2>
          <span className="text-sm text-gray-500">{completed}/{total} complete</span>
        </div>
        <div className="w-full h-2 bg-gray-100 rounded-full overflow-hidden">
          <div
            className="h-full bg-blue-600 rounded-full transition-all duration-500"
            style={{ width: `${(completed / total) * 100}%` }}
          />
        </div>
      </div>

      {allDone ? (
        <div className="px-6 py-8 text-center">
          <CheckCircle2 className="w-12 h-12 text-green-500 mx-auto mb-3" />
          <h3 className="text-lg font-semibold text-gray-900">Setup complete!</h3>
          <p className="text-gray-500 text-sm mt-1">Your HOA portal is fully configured.</p>
        </div>
      ) : (
        <ul className="divide-y divide-gray-50">
          {sorted.map(item => (
            <li key={item.step_key}>
              <Link
                href={`/kevin/admin/settings?step=${item.step_key}`}
                className="flex items-center gap-4 px-6 py-4 hover:bg-gray-50 transition-colors group"
              >
                {item.completed ? (
                  <CheckCircle2 className="w-5 h-5 text-green-500 flex-shrink-0" />
                ) : (
                  <Circle className="w-5 h-5 text-gray-300 flex-shrink-0" />
                )}
                <span className={`flex-1 text-sm font-medium ${item.completed ? 'text-gray-400 line-through' : 'text-gray-700'}`}>
                  {item.step_label}
                </span>
                {!item.completed && (
                  <ChevronRight className="w-4 h-4 text-gray-400 group-hover:text-gray-600 transition-colors" />
                )}
              </Link>
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}
