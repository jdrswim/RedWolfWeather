'use client'

import { useState } from 'react'
import { createClient } from '@/lib/supabaseClient'
import { Upload, FileText } from 'lucide-react'
import { useRouter } from 'next/navigation'

export default function DocumentUpload() {
  const [title, setTitle] = useState('')
  const [file, setFile] = useState<File | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [success, setSuccess] = useState(false)
  const router = useRouter()
  const supabase = createClient()

  async function handleUpload(e: React.FormEvent) {
    e.preventDefault()
    if (!file || !title) return
    setLoading(true)
    setError('')
    setSuccess(false)

    const { data: { user } } = await supabase.auth.getUser()
    if (!user) { setError('Not authenticated'); setLoading(false); return }

    const fileExt = file.name.split('.').pop()
    const filePath = `${Date.now()}-${file.name.replace(/[^a-zA-Z0-9.-]/g, '_')}`

    const { data: uploadData, error: uploadError } = await supabase.storage
      .from('hoa-documents')
      .upload(filePath, file, { cacheControl: '3600', upsert: false })

    if (uploadError) { setError(uploadError.message); setLoading(false); return }

    const { data: { publicUrl } } = supabase.storage
      .from('hoa-documents')
      .getPublicUrl(filePath)

    const { error: dbError } = await supabase.from('documents').insert({
      title,
      file_url: publicUrl,
      file_name: file.name,
      file_size: file.size,
      uploaded_by: user.id,
      is_public: true,
    })

    if (dbError) { setError(dbError.message); setLoading(false); return }

    setSuccess(true)
    setTitle('')
    setFile(null)
    router.refresh()
    setLoading(false)
  }

  return (
    <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6">
      <h2 className="font-semibold text-gray-900 mb-5">Upload Document</h2>
      {error && <div className="mb-4 p-3 bg-red-50 border border-red-200 rounded-lg text-red-700 text-sm">{error}</div>}
      {success && <div className="mb-4 p-3 bg-green-50 border border-green-200 rounded-lg text-green-700 text-sm">Document uploaded successfully!</div>}
      <form onSubmit={handleUpload} className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">Document Title</label>
          <input type="text" value={title} onChange={e => setTitle(e.target.value)} required
            className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500"
            placeholder="HOA Bylaws 2024" />
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1.5">File</label>
          <input type="file" onChange={e => setFile(e.target.files?.[0] || null)} required
            accept=".pdf,.doc,.docx,.txt,.jpg,.png"
            className="w-full px-3 py-2 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm file:mr-3 file:py-1 file:px-3 file:rounded-lg file:border-0 file:bg-blue-50 file:text-blue-600 file:text-xs file:font-medium" />
        </div>
        <div className="flex items-end">
          <button type="submit" disabled={loading || !file || !title}
            className="flex items-center gap-2 px-6 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-300 text-white text-sm font-semibold rounded-xl transition-colors w-full justify-center">
            <Upload size={16} />
            {loading ? 'Uploading...' : 'Upload'}
          </button>
        </div>
      </form>
    </div>
  )
}
