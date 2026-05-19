import { redirect } from 'next/navigation'
import { createClient } from '@/lib/supabaseServer'
import OwnerSidebar from '@/components/layout/OwnerSidebar'
import PageHeader from '@/components/layout/PageHeader'
import { formatDate } from '@/lib/utils'
import { FileText, Download, File } from 'lucide-react'
import type { Document } from '@/types'

export default async function OwnerDocumentsPage() {
  const supabase = await createClient()
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) redirect('/login')

  const [{ data: profile }, { data: docs }, { data: settings }] = await Promise.all([
    supabase.from('profiles').select('*').eq('id', user.id).single(),
    supabase.from('documents').select('*').order('created_at', { ascending: false }),
    supabase.from('hoa_settings').select('hoa_name').single(),
  ])

  const categories = [...new Set(docs?.map((d: any) => d.category).filter(Boolean) as string[])]

  return (
    <div className="flex min-h-screen bg-gray-50">
      <OwnerSidebar
        userName={profile?.name}
        unitNumber={profile?.unit_number ?? undefined}
        hoaName={settings?.hoa_name}
      />
      <main className="flex-1 ml-64 p-8">
        <div className="max-w-5xl mx-auto">
          <PageHeader
            title="Documents"
            description="Access HOA rules, bylaws, and community documents"
          />

          {!docs?.length ? (
            <div className="bg-white rounded-xl border border-gray-100 shadow-sm p-16 text-center">
              <File className="h-12 w-12 text-gray-200 mx-auto mb-4" />
              <p className="font-medium text-gray-500">No documents available yet</p>
              <p className="text-sm text-gray-400 mt-1">Your HOA admin will upload documents here</p>
            </div>
          ) : (
            <>
              {categories.map((cat) => (
                <div key={cat} className="mb-8">
                  <h2 className="text-sm font-semibold text-gray-500 uppercase tracking-wider mb-3">{cat}</h2>
                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                    {docs.filter((d: any) => d.category === cat).map((doc: Document) => (
                      <a
                        key={doc.id}
                        href={doc.file_url}
                        target="_blank"
                        rel="noreferrer"
                        className="bg-white rounded-xl border border-gray-100 shadow-sm p-5 hover:border-blue-100 hover:shadow-md transition-all group"
                      >
                        <div className="flex items-start justify-between mb-3">
                          <div className="h-10 w-10 bg-blue-50 rounded-lg flex items-center justify-center">
                            <FileText className="h-5 w-5 text-blue-600" />
                          </div>
                          <Download className="h-4 w-4 text-gray-300 group-hover:text-blue-500 transition-colors" />
                        </div>
                        <p className="text-sm font-semibold text-gray-900 mb-1 truncate">{doc.title}</p>
                        <p className="text-xs text-gray-400 truncate">{doc.file_name}</p>
                        <p className="text-xs text-gray-400 mt-2">{formatDate(doc.created_at)}</p>
                      </a>
                    ))}
                  </div>
                </div>
              ))}

              {/* Uncategorized */}
              {docs.filter((d: any) => !d.category).length > 0 && (
                <div>
                  <h2 className="text-sm font-semibold text-gray-500 uppercase tracking-wider mb-3">Other</h2>
                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                    {docs.filter((d: any) => !d.category).map((doc: Document) => (
                      <a
                        key={doc.id}
                        href={doc.file_url}
                        target="_blank"
                        rel="noreferrer"
                        className="bg-white rounded-xl border border-gray-100 shadow-sm p-5 hover:border-blue-100 hover:shadow-md transition-all group"
                      >
                        <div className="flex items-start justify-between mb-3">
                          <div className="h-10 w-10 bg-blue-50 rounded-lg flex items-center justify-center">
                            <FileText className="h-5 w-5 text-blue-600" />
                          </div>
                          <Download className="h-4 w-4 text-gray-300 group-hover:text-blue-500 transition-colors" />
                        </div>
                        <p className="text-sm font-semibold text-gray-900 mb-1 truncate">{doc.title}</p>
                        <p className="text-xs text-gray-400 truncate">{doc.file_name}</p>
                        <p className="text-xs text-gray-400 mt-2">{formatDate(doc.created_at)}</p>
                      </a>
                    ))}
                  </div>
                </div>
              )}
            </>
          )}
        </div>
      </main>
    </div>
  )
}
