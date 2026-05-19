import Link from 'next/link'
import {
  Building2,
  CreditCard,
  MessageSquare,
  FileText,
  BarChart3,
  Shield,
  ArrowRight,
  CheckCircle2,
} from 'lucide-react'

export default function LandingPage() {
  return (
    <div className="min-h-screen bg-white">
      {/* Nav */}
      <nav className="border-b border-gray-100 px-6 py-4">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Building2 className="h-7 w-7 text-blue-600" />
            <span className="text-xl font-bold text-gray-900">HOA Manager</span>
          </div>
          <div className="flex items-center gap-4">
            <Link
              href="/login"
              className="text-sm font-medium text-gray-600 hover:text-gray-900 transition-colors"
            >
              Sign in
            </Link>
            <Link
              href="/signup"
              className="bg-blue-600 text-white text-sm font-medium px-4 py-2 rounded-lg hover:bg-blue-700 transition-colors"
            >
              Get started free
            </Link>
          </div>
        </div>
      </nav>

      {/* Hero */}
      <section className="px-6 pt-20 pb-24 bg-gradient-to-b from-blue-50 to-white">
        <div className="max-w-4xl mx-auto text-center">
          <div className="inline-flex items-center gap-2 bg-blue-100 text-blue-700 text-sm font-medium px-3 py-1 rounded-full mb-6">
            <Shield className="h-4 w-4" />
            Built for self-managed HOAs
          </div>
          <h1 className="text-5xl font-bold text-gray-900 mb-6 leading-tight">
            Modern HOA management,{' '}
            <span className="text-blue-600">without the complexity</span>
          </h1>
          <p className="text-xl text-gray-600 mb-10 max-w-2xl mx-auto">
            Collect dues online, manage owners, track expenses, and communicate with your
            community — all in one place. No accounting degree required.
          </p>
          <div className="flex flex-col sm:flex-row gap-4 justify-center">
            <Link
              href="/signup"
              className="inline-flex items-center gap-2 bg-blue-600 text-white font-semibold px-8 py-3.5 rounded-xl hover:bg-blue-700 transition-colors shadow-lg shadow-blue-200"
            >
              Start free setup
              <ArrowRight className="h-5 w-5" />
            </Link>
            <Link
              href="/login"
              className="inline-flex items-center gap-2 bg-white text-gray-700 font-semibold px-8 py-3.5 rounded-xl border border-gray-200 hover:border-gray-300 hover:bg-gray-50 transition-colors"
            >
              Sign in to your account
            </Link>
          </div>
        </div>
      </section>

      {/* Features */}
      <section className="px-6 py-24 max-w-7xl mx-auto">
        <div className="text-center mb-16">
          <h2 className="text-3xl font-bold text-gray-900 mb-4">
            Everything your HOA needs
          </h2>
          <p className="text-gray-600 max-w-xl mx-auto">
            From dues collection to document storage, manage your entire community from a
            single dashboard.
          </p>
        </div>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
          {features.map((f) => (
            <div
              key={f.title}
              className="p-6 rounded-2xl border border-gray-100 hover:border-blue-100 hover:shadow-lg hover:shadow-blue-50 transition-all"
            >
              <div className="h-12 w-12 bg-blue-50 rounded-xl flex items-center justify-center mb-4">
                <f.icon className="h-6 w-6 text-blue-600" />
              </div>
              <h3 className="text-lg font-semibold text-gray-900 mb-2">{f.title}</h3>
              <p className="text-gray-600 text-sm leading-relaxed">{f.description}</p>
            </div>
          ))}
        </div>
      </section>

      {/* Checklist */}
      <section className="px-6 py-24 bg-blue-600">
        <div className="max-w-3xl mx-auto text-center">
          <h2 className="text-3xl font-bold text-white mb-4">
            Set up in minutes, not days
          </h2>
          <p className="text-blue-100 mb-10">
            Our guided onboarding walks you through every step to get your HOA running.
          </p>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-left mb-10">
            {setupSteps.map((step) => (
              <div key={step} className="flex items-center gap-3 bg-blue-500 rounded-xl p-4">
                <CheckCircle2 className="h-5 w-5 text-blue-200 flex-shrink-0" />
                <span className="text-white text-sm font-medium">{step}</span>
              </div>
            ))}
          </div>
          <Link
            href="/signup"
            className="inline-flex items-center gap-2 bg-white text-blue-600 font-semibold px-8 py-3.5 rounded-xl hover:bg-blue-50 transition-colors"
          >
            Get started — it&apos;s free
            <ArrowRight className="h-5 w-5" />
          </Link>
        </div>
      </section>

      {/* Footer */}
      <footer className="border-t border-gray-100 px-6 py-10">
        <div className="max-w-7xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-2">
            <Building2 className="h-5 w-5 text-blue-600" />
            <span className="font-semibold text-gray-900">HOA Manager</span>
          </div>
          <p className="text-sm text-gray-500">
            © {new Date().getFullYear()} HOA Manager. Built with Next.js + Supabase.
          </p>
        </div>
      </footer>
    </div>
  )
}

const features = [
  {
    icon: CreditCard,
    title: 'Online Dues Collection',
    description:
      'Accept card and ACH payments via Stripe. Automatically track balances and send reminders.',
  },
  {
    icon: Building2,
    title: 'Owner Management',
    description:
      'Manage all unit owners, their contact info, and account status in one place.',
  },
  {
    icon: MessageSquare,
    title: 'Community Messaging',
    description:
      'Send broadcasts to all owners or direct messages. Real-time delivery powered by Supabase.',
  },
  {
    icon: BarChart3,
    title: 'Financial Dashboard',
    description:
      'Track income, expenses, and reserve funds. Export reports for your annual meeting.',
  },
  {
    icon: FileText,
    title: 'Document Storage',
    description:
      'Store bylaws, meeting minutes, and community rules. Owners can access approved documents anytime.',
  },
  {
    icon: Shield,
    title: 'Role-Based Access',
    description:
      'Admins see everything; owners only see their own data. Powered by Supabase Row Level Security.',
  },
]

const setupSteps = [
  'HOA organization & unit setup',
  'Connect your Stripe account',
  'Configure payment rules & due dates',
  'Add owners and send invites',
  'Upload founding documents',
  'Start collecting dues online',
]
