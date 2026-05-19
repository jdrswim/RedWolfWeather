'use client'

import { useEffect, useState, useCallback, useRef } from 'react'
import { createClient } from '@/lib/supabaseClient'
import AdminSidebar from '@/components/layout/AdminSidebar'
import PageHeader from '@/components/layout/PageHeader'
import Button from '@/components/ui/Button'
import { formatDateTime, getInitials } from '@/lib/utils'
import { Send, Users, User, Radio } from 'lucide-react'
import type { Message, Profile } from '@/types'

export default function AdminMessagesPage() {
  const [messages, setMessages] = useState<Message[]>([])
  const [owners, setOwners] = useState<Profile[]>([])
  const [selectedOwner, setSelectedOwner] = useState<Profile | null>(null)
  const [isBroadcast, setIsBroadcast] = useState(false)
  const [content, setContent] = useState('')
  const [sending, setSending] = useState(false)
  const [currentUser, setCurrentUser] = useState<Profile | null>(null)
  const [hoaName, setHoaName] = useState('')
  const messagesEndRef = useRef<HTMLDivElement>(null)
  const supabase = createClient()

  const fetchMessages = useCallback(async () => {
    if (!currentUser) return
    let query = supabase
      .from('messages')
      .select('*, sender:profiles!sender_id(*), recipient:profiles!recipient_id(*)')
      .order('created_at', { ascending: true })

    if (isBroadcast) {
      query = query.eq('is_broadcast', true)
    } else if (selectedOwner) {
      query = query.or(
        `and(sender_id.eq.${currentUser.id},recipient_id.eq.${selectedOwner.id}),and(sender_id.eq.${selectedOwner.id},recipient_id.eq.${currentUser.id})`
      )
    } else {
      query = query.limit(0)
    }

    const { data } = await query
    setMessages((data as Message[]) ?? [])
  }, [currentUser, isBroadcast, selectedOwner, supabase])

  useEffect(() => {
    supabase.auth.getUser().then(({ data: { user } }) => {
      if (user) {
        supabase.from('profiles').select('*').eq('id', user.id).single().then(({ data }) => {
          setCurrentUser(data)
        })
      }
    })
    supabase.from('profiles').select('*').eq('role', 'owner').then(({ data }) => setOwners(data ?? []))
    supabase.from('hoa_settings').select('hoa_name').single().then(({ data }) => { if (data) setHoaName(data.hoa_name) })
  }, [supabase])

  useEffect(() => {
    fetchMessages()
  }, [fetchMessages])

  // Realtime subscription
  useEffect(() => {
    const channel = supabase
      .channel('admin-messages')
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
    if (!content.trim() || !currentUser) return
    setSending(true)

    await supabase.from('messages').insert({
      sender_id: currentUser.id,
      recipient_id: isBroadcast ? null : selectedOwner?.id ?? null,
      content: content.trim(),
      is_broadcast: isBroadcast,
      read_status: 'unread',
    })

    setContent('')
    setSending(false)
    fetchMessages()
  }

  return (
    <div className="flex min-h-screen bg-gray-50">
      <AdminSidebar hoaName={hoaName} />
      <main className="flex-1 ml-64 flex">
        {/* Owner list sidebar */}
        <div className="w-72 bg-white border-r border-gray-100 flex flex-col">
          <div className="px-4 py-4 border-b border-gray-100">
            <h2 className="font-semibold text-gray-900">Messages</h2>
          </div>
          <div className="flex-1 overflow-y-auto">
            {/* Broadcast option */}
            <button
              onClick={() => { setIsBroadcast(true); setSelectedOwner(null) }}
              className={`w-full flex items-center gap-3 px-4 py-3.5 text-left hover:bg-gray-50 transition-colors border-b border-gray-50 ${
                isBroadcast ? 'bg-blue-50' : ''
              }`}
            >
              <div className="h-9 w-9 bg-orange-100 rounded-full flex items-center justify-center flex-shrink-0">
                <Radio className="h-4 w-4 text-orange-600" />
              </div>
              <div>
                <p className={`text-sm font-medium ${isBroadcast ? 'text-blue-700' : 'text-gray-900'}`}>
                  Broadcast to all
                </p>
                <p className="text-xs text-gray-400">Send to all owners</p>
              </div>
            </button>

            {/* Owner list */}
            {owners.map((owner) => (
              <button
                key={owner.id}
                onClick={() => { setSelectedOwner(owner); setIsBroadcast(false) }}
                className={`w-full flex items-center gap-3 px-4 py-3.5 text-left hover:bg-gray-50 transition-colors border-b border-gray-50 ${
                  selectedOwner?.id === owner.id && !isBroadcast ? 'bg-blue-50' : ''
                }`}
              >
                <div className="h-9 w-9 bg-blue-100 rounded-full flex items-center justify-center flex-shrink-0">
                  <span className="text-xs font-semibold text-blue-700">
                    {getInitials(owner.name || owner.email || '?')}
                  </span>
                </div>
                <div className="min-w-0">
                  <p className={`text-sm font-medium truncate ${selectedOwner?.id === owner.id && !isBroadcast ? 'text-blue-700' : 'text-gray-900'}`}>
                    {owner.name}
                  </p>
                  <p className="text-xs text-gray-400">
                    {owner.unit_number ? `Unit ${owner.unit_number}` : owner.email}
                  </p>
                </div>
              </button>
            ))}
          </div>
        </div>

        {/* Chat area */}
        <div className="flex-1 flex flex-col">
          {!selectedOwner && !isBroadcast ? (
            <div className="flex-1 flex items-center justify-center text-center">
              <div>
                <Users className="h-10 w-10 text-gray-200 mx-auto mb-3" />
                <p className="text-gray-500 font-medium">Select an owner or broadcast</p>
                <p className="text-gray-400 text-sm mt-1">Choose from the left to start messaging</p>
              </div>
            </div>
          ) : (
            <>
              {/* Header */}
              <div className="bg-white border-b border-gray-100 px-6 py-4">
                <div className="flex items-center gap-3">
                  {isBroadcast ? (
                    <>
                      <div className="h-8 w-8 bg-orange-100 rounded-full flex items-center justify-center">
                        <Radio className="h-4 w-4 text-orange-600" />
                      </div>
                      <div>
                        <p className="font-semibold text-gray-900">Broadcast</p>
                        <p className="text-xs text-gray-400">Message visible to all owners</p>
                      </div>
                    </>
                  ) : (
                    <>
                      <div className="h-8 w-8 bg-blue-100 rounded-full flex items-center justify-center">
                        <span className="text-xs font-semibold text-blue-700">
                          {getInitials(selectedOwner!.name || '?')}
                        </span>
                      </div>
                      <div>
                        <p className="font-semibold text-gray-900">{selectedOwner!.name}</p>
                        <p className="text-xs text-gray-400">
                          {selectedOwner!.unit_number ? `Unit ${selectedOwner!.unit_number}` : selectedOwner!.email}
                        </p>
                      </div>
                    </>
                  )}
                </div>
              </div>

              {/* Messages */}
              <div className="flex-1 overflow-y-auto p-6 space-y-4">
                {messages.length === 0 && (
                  <div className="text-center text-gray-400 text-sm mt-10">
                    No messages yet. Start the conversation!
                  </div>
                )}
                {messages.map((msg) => {
                  const isMe = msg.sender_id === currentUser?.id
                  return (
                    <div key={msg.id} className={`flex ${isMe ? 'justify-end' : 'justify-start'}`}>
                      <div className={`max-w-sm ${isMe ? 'items-end' : 'items-start'} flex flex-col gap-1`}>
                        <div
                          className={`px-4 py-2.5 rounded-2xl text-sm ${
                            isMe
                              ? 'bg-blue-600 text-white rounded-br-sm'
                              : 'bg-white border border-gray-100 text-gray-900 rounded-bl-sm shadow-sm'
                          }`}
                        >
                          {msg.content}
                        </div>
                        <p className="text-xs text-gray-400">
                          {isMe ? 'You' : (msg.sender as Profile)?.name || 'Unknown'} · {formatDateTime(msg.created_at)}
                        </p>
                      </div>
                    </div>
                  )
                })}
                <div ref={messagesEndRef} />
              </div>

              {/* Input */}
              <form onSubmit={sendMessage} className="bg-white border-t border-gray-100 px-6 py-4">
                <div className="flex gap-3">
                  <input
                    type="text"
                    value={content}
                    onChange={(e) => setContent(e.target.value)}
                    placeholder={isBroadcast ? 'Broadcast message to all owners…' : `Message ${selectedOwner?.name}…`}
                    className="flex-1 px-4 py-2.5 border border-gray-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                  />
                  <Button type="submit" loading={sending} disabled={!content.trim()}>
                    <Send className="h-4 w-4" />
                    Send
                  </Button>
                </div>
              </form>
            </>
          )}
        </div>
      </main>
    </div>
  )
}
