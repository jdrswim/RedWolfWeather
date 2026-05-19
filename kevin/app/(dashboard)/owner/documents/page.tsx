import { createServerSupabaseClient } from '@/lib/supabaseServer'
import { redirect } from 'next/navigation'
import { FileText, Download } from 'lucide-react'
import { format } from 'date-fns'

export default async function OwnerDocumentsPage() {
  const supabase = await createServerSupabaseClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) redirect('/kevin/login')

  const { data: documents } = await supabase
    .from('documents')
    .select('*')
    .eq('is_public', true)
    .order('created_at', { ascending: false })

  return (
    <div className="max-w-4xl mx-auto space-y-6">
      <div>
        <h1 className="text-2xl font-bold text-gray-900">HOA Documents</h1>
        <p className="text-gray-500 text-sm mt-1">Download bylaws, rules, and official HOA documents</p>
      </div>

      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
        {!documents || documents.length === 0 ? (
          <div className="text-center py-16 text-gray-400">
            <FileText className="w-12 h-12 mx-auto mb-3 opacity-40" />
            <p className="font-medium">No documents available</p>
            <p className="text-sm mt-1">Check back later for HOA documents</p>
          </div>
        ) : (
          <ul className="divide-y divide-gray-50">
            {documents.map(doc => (
              <li key={doc.id} className="flex items-center gap-4 px-6 py-4 hover:bg-gray-50 transition-colors">
                <div className="p-2.5 bg-blue-50 rounded-xl">
                  <FileText className="w-5 h-5 text-blue-600" />
                </div>
                <div className="flex-1 min-w-0">
                  <p className="font-medium text-gray-900 truncate">{doc.title}</p>
                  <p className="text-xs text-gray-400 mt-0.5">
                    {doc.file_name} • Added {format(new Date(doc.created_at), 'MMM d, yyyy')}
                  </p>
                </div>
                <a
                  href={doc.file_url}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="flex items-center gap-1.5 px-3 py-1.5 text-sm font-medium text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                >
                  <Download size={14} />
                  Download
                </a>
              </li>
            ))}
          </ul>
        )}
      </div>
    </div>
  )
}
