import { createServerSupabaseClient } from '@/lib/supabaseServer'
import { redirect } from 'next/navigation'
import DocumentUpload from './DocumentUpload'
import { FileText, Download, Trash2 } from 'lucide-react'
import { format } from 'date-fns'

export default async function AdminDocumentsPage() {
  const supabase = await createServerSupabaseClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) redirect('/kevin/login')

  const { data: profile } = await supabase.from('profiles').select('role').eq('id', user.id).single()
  if (profile?.role !== 'admin') redirect('/kevin/owner')

  const { data: documents } = await supabase
    .from('documents')
    .select('*, uploader:profiles!uploaded_by(name)')
    .order('created_at', { ascending: false })

  return (
    <div className="max-w-5xl mx-auto space-y-6">
      <div>
        <h1 className="text-2xl font-bold text-gray-900">Documents</h1>
        <p className="text-gray-500 text-sm mt-1">Upload and manage HOA documents</p>
      </div>

      <DocumentUpload />

      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="px-6 py-4 border-b border-gray-100">
          <h2 className="font-semibold text-gray-900">All Documents ({documents?.length || 0})</h2>
        </div>
        {!documents || documents.length === 0 ? (
          <div className="text-center py-16 text-gray-400">
            <FileText className="w-12 h-12 mx-auto mb-3 opacity-40" />
            <p className="font-medium">No documents uploaded</p>
            <p className="text-sm mt-1">Upload bylaws, rules, and other HOA documents above</p>
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
                    Uploaded by {(doc.uploader as any)?.name || 'Admin'} • {format(new Date(doc.created_at), 'MMM d, yyyy')}
                  </p>
                </div>
                <div className="flex items-center gap-2">
                  <a
                    href={doc.file_url}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="p-2 text-gray-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                    title="Download"
                  >
                    <Download size={16} />
                  </a>
                </div>
              </li>
            ))}
          </ul>
        )}
      </div>
    </div>
  )
}
