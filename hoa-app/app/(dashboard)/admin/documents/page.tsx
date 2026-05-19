'use client'

import { useEffect, useState, useCallback, useRef } from 'react'
import { createClient } from '@/lib/supabaseClient'
import AdminSidebar from '@/components/layout/AdminSidebar'
import PageHeader from '@/components/layout/PageHeader'
import Button from '@/components/ui/Button'
import { formatDate } from '@/lib/utils'
import { Upload, FileText, Download, Trash2, File } from 'lucide-react'
import type { Document } from '@/types'

const CATEGORIES = ['Bylaws', 'Rules & Regulations', 'Meeting Minutes', 'Financial Reports', 'Insurance', 'Maintenance', 'Other']

export default function DocumentsPage() {
  const [docs, setDocs] = useState<Document[]>([])
  const [loading, setLoading] = useState(true)
  const [uploading, setUploading] = useState(false)
  const [categoryFilter, setCategoryFilter] = useState('all')
  const [hoaName, setHoaName] = useState('')
  const [error, setError] = useState('')
  const [category, setCategory] = useState('Other')
  const [title, setTitle] = useState('')
  const fileRef = useRef<HTMLInputElement>(null)
  const supabase = createClient()

  const fetchDocs = useCallback(async () => {
    const { data } = await supabase
      .from('documents')
      .select('*, uploader:profiles(name)')
      .order('created_at', { ascending: false })
    setDocs((data as Document[]) ?? [])
    setLoading(false)
  }, [supabase])

  useEffect(() => {
    fetchDocs()
    supabase.from('hoa_settings').select('hoa_name').single().then(({ data }) => { if (data) setHoaName(data.hoa_name) })
  }, [fetchDocs, supabase])

  async function handleUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file) return
    if (!title) { setError('Please enter a document title first'); return }
    setUploading(true)
    setError('')

    const { data: { user } } = await supabase.auth.getUser()
    const ext = file.name.split('.').pop()
    const path = `${user!.id}/${Date.now()}.${ext}`

    const { error: uploadErr } = await supabase.storage
      .from('hoa-documents')
      .upload(path, file)

    if (uploadErr) { setError(uploadErr.message); setUploading(false); return }

    const { data: { publicUrl } } = supabase.storage
      .from('hoa-documents')
      .getPublicUrl(path)

    await supabase.from('documents').insert({
      title,
      file_url: publicUrl,
      file_name: file.name,
      file_size: file.size,
      uploaded_by: user!.id,
      category,
    })

    setTitle('')
    if (fileRef.current) fileRef.current.value = ''
    fetchDocs()
    setUploading(false)
  }

  async function deleteDoc(doc: Document) {
    if (!confirm(`Delete "${doc.title}"?`)) return
    await supabase.from('documents').delete().eq('id', doc.id)
    fetchDocs()
  }

  const filtered = categoryFilter === 'all' ? docs : docs.filter((d) => d.category === categoryFilter)

  return (
    <div className="flex min-h-screen bg-gray-50">
      <AdminSidebar hoaName={hoaName} />
      <main className="flex-1 ml-64 p-8">
        <div className="max-w-5xl mx-auto">
          <PageHeader
            title="Documents"
            description="Store and manage HOA documents for all owners"
          />

          {/* Upload panel */}
          <div className="bg-white rounded-xl border border-gray-100 shadow-sm p-6 mb-6">
            <h3 className="font-semibold text-gray-900 mb-4">Upload document</h3>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-4">
              <input
                type="text"
                placeholder="Document title *"
                value={title}
                onChange={(e) => setTitle(e.target.value)}
                className="px-3.5 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 sm:col-span-2"
              />
              <select
                value={category}
                onChange={(e) => setCategory(e.target.value)}
                className="px-3.5 py-2.5 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
              >
                {CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>
            {error && <p className="text-sm text-red-600 mb-3">{error}</p>}
            <div className="flex items-center gap-3">
              <label className="cursor-pointer">
                <input
                  ref={fileRef}
                  type="file"
                  className="hidden"
                  onChange={handleUpload}
                  accept=".pdf,.doc,.docx,.xls,.xlsx,.png,.jpg,.jpeg"
                />
                <Button as="span" loading={uploading} disabled={!title || uploading}>
                  <Upload className="h-4 w-4" />
                  {uploading ? 'Uploading…' : 'Choose file & upload'}
                </Button>
              </label>
              <p className="text-xs text-gray-400">PDF, Word, Excel, or images up to 50MB</p>
            </div>
          </div>

          {/* Filter */}
          <div className="flex gap-2 mb-5 flex-wrap">
            {['all', ...CATEGORIES].map((c) => (
              <button
                key={c}
                onClick={() => setCategoryFilter(c)}
                className={`text-xs font-medium px-3 py-1.5 rounded-full transition-colors ${
                  categoryFilter === c
                    ? 'bg-blue-600 text-white'
                    : 'bg-white border border-gray-200 text-gray-600 hover:border-gray-300'
                }`}
              >
                {c === 'all' ? 'All' : c}
              </button>
            ))}
          </div>

          {/* Doc grid */}
          {loading ? (
            <div className="text-center text-gray-400 py-12">Loading…</div>
          ) : filtered.length === 0 ? (
            <div className="text-center py-12">
              <File className="h-10 w-10 text-gray-200 mx-auto mb-3" />
              <p className="text-gray-500">No documents uploaded yet</p>
            </div>
          ) : (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {filtered.map((doc) => (
                <div key={doc.id} className="bg-white rounded-xl border border-gray-100 shadow-sm p-5 hover:border-blue-100 transition-colors group">
                  <div className="flex items-start justify-between mb-3">
                    <div className="h-10 w-10 bg-blue-50 rounded-lg flex items-center justify-center">
                      <FileText className="h-5 w-5 text-blue-600" />
                    </div>
                    <div className="flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                      <a href={doc.file_url} target="_blank" rel="noreferrer">
                        <Button variant="ghost" size="sm">
                          <Download className="h-3.5 w-3.5" />
                        </Button>
                      </a>
                      <Button variant="ghost" size="sm" onClick={() => deleteDoc(doc)}>
                        <Trash2 className="h-3.5 w-3.5 text-red-400" />
                      </Button>
                    </div>
                  </div>
                  <p className="text-sm font-semibold text-gray-900 mb-1 truncate">{doc.title}</p>
                  <p className="text-xs text-gray-400 mb-2 truncate">{doc.file_name}</p>
                  <div className="flex items-center justify-between">
                    <span className="text-xs bg-gray-100 text-gray-600 px-2 py-0.5 rounded-full">{doc.category}</span>
                    <span className="text-xs text-gray-400">{formatDate(doc.created_at)}</span>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </main>
    </div>
  )
}
