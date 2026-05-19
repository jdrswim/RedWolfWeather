'use client'

import { useState, useEffect, useRef } from 'react'
import { createClient } from '@/lib/supabaseClient'
import { Send, Users, User } from 'lucide-react'
import { format } from 'date-fns'
import { useRouter } from 'next/navigation'

interface Recipient { id: string; name: string; unit_number?: string | null }
interface CurrentUser { id: string; name: string; role: string }
interface Message {
  id: string
  sender_id: string
  recipient_id: string | null
  content: string
  is_broadcast: boolean
  read_status: boolean
  created_at: string
  sender?: { name: string } | null
  recipient?: { name: string } | null
}

interface Props {
  currentUser: CurrentUser
  recipients: Recipient[]
  messages: Message[]
}

export default function MessagingPanel({ currentUser, recipients, messages: initialMessages }: Props) {
  const [messages, setMessages] = useState(initialMessages)
  const [content, setContent] = useState('')
  const [recipientId, setRecipientId] = useState<string | null>(null)
  const [isBroadcast, setIsBroadcast] = useState(false)
  const [sending, setSending] = useState(false)
  const supabase = createClient()
  const router = useRouter()
  const messagesEndRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const channel = supabase
      .channel('messages')
      .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'messages' }, () => {
        router.refresh()
      })
      .subscribe()
    return () => { supabase.removeChannel(channel) }
  }, [])

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [messages])

  async function sendMessage(e: React.FormEvent) {
    e.preventDefault()
    if (!content.trim()) return
    setSending(true)

    const { error } = await supabase.from('messages').insert({
      sender_id: currentUser.id,
      recipient_id: isBroadcast ? null : recipientId,
      content: content.trim(),
      is_broadcast: isBroadcast,
    })

    if (!error) {
      setContent('')
      router.refresh()
    }
    setSending(false)
  }

  const filteredMessages = messages.filter(m => {
    if (m.is_broadcast) return true
    return m.sender_id === currentUser.id || m.recipient_id === currentUser.id
  })

  return (
    <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
      {/* Message list */}
      <div className="lg:col-span-2 bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden flex flex-col" style={{ height: '60vh' }}>
        <div className="px-6 py-4 border-b border-gray-100">
          <h2 className="font-semibold text-gray-900">Conversation</h2>
        </div>
        <div className="flex-1 overflow-y-auto p-4 space-y-3">
          {filteredMessages.length === 0 ? (
            <div className="flex items-center justify-center h-full text-gray-400 text-sm">
              No messages yet — start a conversation below
            </div>
          ) : (
            filteredMessages.slice().reverse().map(m => {
              const isOwn = m.sender_id === currentUser.id
              return (
                <div key={m.id} className={`flex ${isOwn ? 'justify-end' : 'justify-start'}`}>
                  <div className={`max-w-xs lg:max-w-md px-4 py-3 rounded-2xl ${
                    isOwn
                      ? 'bg-blue-600 text-white rounded-br-sm'
                      : m.is_broadcast
                      ? 'bg-amber-50 border border-amber-200 text-gray-800 rounded-bl-sm'
                      : 'bg-gray-100 text-gray-800 rounded-bl-sm'
                  }`}>
                    <div className="flex items-center gap-2 mb-1">
                      {m.is_broadcast && <Users size={12} className="text-amber-600" />}
                      <p className={`text-xs font-medium ${isOwn ? 'text-blue-100' : m.is_broadcast ? 'text-amber-600' : 'text-gray-500'}`}>
                        {m.is_broadcast ? 'Broadcast' : (m.sender as any)?.name || 'Unknown'}
                      </p>
                    </div>
                    <p className="text-sm leading-relaxed">{m.content}</p>
                    <p className={`text-xs mt-1 ${isOwn ? 'text-blue-200' : 'text-gray-400'}`}>
                      {format(new Date(m.created_at), 'MMM d, h:mm a')}
                    </p>
                  </div>
                </div>
              )
            })
          )}
          <div ref={messagesEndRef} />
        </div>
      </div>

      {/* Compose */}
      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6 space-y-5">
        <h2 className="font-semibold text-gray-900">Compose</h2>

        {currentUser.role === 'admin' && (
          <div>
            <label className="flex items-center gap-2 cursor-pointer">
              <input type="checkbox" checked={isBroadcast} onChange={e => setIsBroadcast(e.target.checked)}
                className="w-4 h-4 text-blue-600 rounded" />
              <span className="text-sm font-medium text-gray-700 flex items-center gap-1.5">
                <Users size={15} /> Broadcast to all owners
              </span>
            </label>
          </div>
        )}

        {!isBroadcast && (
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">
              {currentUser.role === 'admin' ? 'Send to Owner' : 'Send to Admin'}
            </label>
            <select value={recipientId || ''} onChange={e => setRecipientId(e.target.value || null)}
              className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white text-sm">
              <option value="">
                {currentUser.role === 'admin' ? 'Select owner...' : 'Select admin...'}
              </option>
              {recipients.map(r => (
                <option key={r.id} value={r.id}>{r.name}</option>
              ))}
            </select>
          </div>
        )}

        <form onSubmit={sendMessage} className="space-y-3">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1.5">Message</label>
            <textarea value={content} onChange={e => setContent(e.target.value)} rows={4} required
              className="w-full px-3 py-2.5 border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 resize-none text-sm"
              placeholder={isBroadcast ? 'Broadcast message to all owners...' : 'Type your message...'} />
          </div>
          <button type="submit" disabled={sending || (!isBroadcast && !recipientId)}
            className="w-full flex items-center justify-center gap-2 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-300 text-white text-sm font-semibold rounded-xl transition-colors">
            <Send size={15} />
            {sending ? 'Sending...' : 'Send Message'}
          </button>
        </form>
      </div>
    </div>
  )
}
