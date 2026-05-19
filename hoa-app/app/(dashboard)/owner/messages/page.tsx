'use client'

import { useEffect, useState, useCallback, useRef } from 'react'
import { createClient } from '@/lib/supabaseClient'
import OwnerSidebar from '@/components/layout/OwnerSidebar'
import Button from '@/components/ui/Button'
import { formatDateTime } from '@/lib/utils'
import { Send, MessageSquare, Radio } from 'lucide-react'
import type { Message, Profile, HoaSettings } from '@/types'

export default function OwnerMessagesPage() {
  const [tab, setTab] = useState<'direct' | 'broadcast'>('direct')
  const [messages, setMessages] = useState<Message[]>([])
  const [content, setContent] = useState('')
  const [sending, setSending] = useState(false)
  const [currentUser, setCurrentUser] = useState<Profile | null>(null)
  const [admin, setAdmin] = useState<Profile | null>(null)
  const [settings, setSettings] = useState<HoaSettings | null>(null)
  const messagesEndRef = useRef<HTMLDivElement>(null)
  const supabase = createClient()

  const fetchMessages = useCallback(async () => {
    if (!currentUser || !admin) return

    if (tab === 'broadcast') {
      const { data } = await supabase
        .from('messages')
        .select('*, sender:profiles!sender_id(*)')
        .eq('is_broadcast', true)
        .order('created_at', { ascending: true })
      setMessages((data as Message[]) ?? [])
    } else {
      const { data } = await supabase
        .from('messages')
        .select('*, sender:profiles!sender_id(*)')
        .eq('is_broadcast', false)
        .or(`and(sender_id.eq.${currentUser.id},recipient_id.eq.${admin.id}),and(sender_id.eq.${admin.id},recipient_id.eq.${currentUser.id})`)
        .order('created_at', { ascending: true })
      setMessages((data as Message[]) ?? [])

      // Mark received messages as read
      await supabase
        .from('messages')
        .update({ read_status: 'read' })
        .eq('recipient_id', currentUser.id)
        .eq('sender_id', admin.id)
        .eq('read_status', 'unread')
    }
  }, [currentUser, admin, tab, supabase])

  useEffect(() => {
    supabase.auth.getUser().then(({ data: { user } }) => {
      if (user) {
        supabase.from('profiles').select('*').eq('id', user.id).single().then(({ data }) => setCurrentUser(data))
      }
    })
    supabase.from('profiles').select('*').eq('role', 'admin').limit(1).single().then(({ data }) => setAdmin(data))
    supabase.from('hoa_settings').select('*').single().then(({ data }) => setSettings(data))
  }, [supabase])

  useEffect(() => { fetchMessages() }, [fetchMessages])

  useEffect(() => {
    const channel = supabase
      .channel('owner-messages')
      .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'messages' }, () => {
        fetchMessages()
      })
      .subscribe()
    return () => { supabase.removeChannel(channel) }
  }, [fetchMessages, supabase])

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [messages])

  async function sendMessage(e: React.FormEvent) {
    e.preventDefault()
    if (!content.trim() || !currentUser || !admin) return
    setSending(true)

    await supabase.from('messages').insert({
      sender_id: currentUser.id,
      recipient_id: admin.id,
      content: content.trim(),
      is_broadcast: false,
      read_status: 'unread',
    })

    setContent('')
    setSending(false)
    fetchMessages()
  }

  return (
    <div className="flex min-h-screen bg-gray-50">
      <OwnerSidebar
        userName={currentUser?.name}
        unitNumber={currentUser?.unit_number ?? undefined}
        hoaName={settings?.hoa_name}
      />
      <main className="flex-1 ml-64 flex flex-col max-h-screen">
        {/* Tabs */}
        <div className="bg-white border-b border-gray-100 px-6 py-0 flex gap-1 pt-4">
          <button
            onClick={() => setTab('direct')}
            className={`flex items-center gap-2 px-4 py-3 text-sm font-medium border-b-2 transition-colors ${
              tab === 'direct'
                ? 'border-blue-600 text-blue-700'
                : 'border-transparent text-gray-500 hover:text-gray-700'
            }`}
          >
            <MessageSquare className="h-4 w-4" />
            Direct messages
          </button>
          <button
            onClick={() => setTab('broadcast')}
            className={`flex items-center gap-2 px-4 py-3 text-sm font-medium border-b-2 transition-colors ${
              tab === 'broadcast'
                ? 'border-blue-600 text-blue-700'
                : 'border-transparent text-gray-500 hover:text-gray-700'
            }`}
          >
            <Radio className="h-4 w-4" />
            Announcements
          </button>
        </div>

        {/* Messages */}
        <div className="flex-1 overflow-y-auto p-6 space-y-4">
          {messages.length === 0 && (
            <div className="text-center text-gray-400 text-sm mt-10">
              {tab === 'direct'
                ? 'No messages yet. Send a message to your HOA admin.'
                : 'No announcements from your HOA admin yet.'}
            </div>
          )}
          {messages.map((msg) => {
            const isMe = msg.sender_id === currentUser?.id
            return (
              <div key={msg.id} className={`flex ${isMe ? 'justify-end' : 'justify-start'}`}>
                <div className={`max-w-sm flex flex-col gap-1 ${isMe ? 'items-end' : 'items-start'}`}>
                  <div
                    className={`px-4 py-2.5 rounded-2xl text-sm ${
                      isMe
                        ? 'bg-blue-600 text-white rounded-br-sm'
                        : msg.is_broadcast
                        ? 'bg-orange-50 border border-orange-100 text-gray-900 rounded-bl-sm'
                        : 'bg-white border border-gray-100 text-gray-900 rounded-bl-sm shadow-sm'
                    }`}
                  >
                    {msg.is_broadcast && !isMe && (
                      <p className="text-xs text-orange-500 font-medium mb-1 flex items-center gap-1">
                        <Radio className="h-3 w-3" />
                        Announcement
                      </p>
                    )}
                    {msg.content}
                  </div>
                  <p className="text-xs text-gray-400">
                    {isMe ? 'You' : (msg.sender as Profile)?.name || 'Admin'} · {formatDateTime(msg.created_at)}
                  </p>
                </div>
              </div>
            )
          })}
          <div ref={messagesEndRef} />
        </div>

        {/* Input — only for direct messages */}
        {tab === 'direct' && (
          <form onSubmit={sendMessage} className="bg-white border-t border-gray-100 px-6 py-4">
            <div className="flex gap-3">
              <input
                type="text"
                value={content}
                onChange={(e) => setContent(e.target.value)}
                placeholder="Message your HOA admin…"
                className="flex-1 px-4 py-2.5 border border-gray-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
              />
              <Button type="submit" loading={sending} disabled={!content.trim()}>
                <Send className="h-4 w-4" />
                Send
              </Button>
            </div>
          </form>
        )}
      </main>
    </div>
  )
}
