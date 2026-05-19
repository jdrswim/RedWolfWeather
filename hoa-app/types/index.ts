export type UserRole = 'owner' | 'admin'
export type DuesStatus = 'paid' | 'unpaid' | 'overdue' | 'partial'
export type PaymentStatus = 'pending' | 'completed' | 'failed' | 'refunded'
export type MessageReadStatus = 'read' | 'unread'

export interface Profile {
  id: string
  name: string
  email: string
  role: UserRole
  unit_number: string | null
  phone: string | null
  created_at: string
}

export interface Due {
  id: string
  owner_id: string
  amount_due: number
  due_date: string
  status: DuesStatus
  balance_remaining: number
  month_year: string
  notes: string | null
  created_at: string
  profiles?: Profile
}

export interface Payment {
  id: string
  owner_id: string
  amount: number
  stripe_session_id: string | null
  stripe_payment_intent_id: string | null
  status: PaymentStatus
  payment_date: string
  due_id: string | null
  created_at: string
  profiles?: Profile
}

export interface Message {
  id: string
  sender_id: string
  recipient_id: string | null
  content: string
  read_status: MessageReadStatus
  is_broadcast: boolean
  created_at: string
  sender?: Profile
  recipient?: Profile
}

export interface Expense {
  id: string
  vendor_name: string
  category: string
  amount: number
  date: string
  notes: string | null
  receipt_url: string | null
  created_by: string
  created_at: string
}

export interface Document {
  id: string
  title: string
  file_url: string
  file_name: string
  file_size: number | null
  uploaded_by: string
  category: string | null
  created_at: string
  uploader?: Profile
}

export interface HoaSettings {
  id: string
  hoa_name: string
  address: string
  city: string
  state: string
  zip: string
  num_units: number
  default_monthly_dues: number | null
  default_due_day: number
  late_fee_amount: number | null
  late_fee_days_grace: number | null
  accounting_start_date: string | null
  payment_methods: string[]
  stripe_configured: boolean
  onboarding_completed: boolean
  created_at: string
  updated_at: string
}

export interface OnboardingStep {
  id: string
  label: string
  description: string
  completed: boolean
  path: string
}
