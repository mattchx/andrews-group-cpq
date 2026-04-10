import { useState } from 'react'
import { sections } from '../lib/questions'
import { calculateRiskScore } from '../lib/scoring'
import { generateRecommendations, type Recommendation } from '../lib/recommendations'
import { exportPdf } from '../lib/exportPdf'
import { exportDocx } from '../lib/exportDocx'

interface ResultsViewProps {
  answers: Record<string, string | string[]>
  onRestart: () => void
  onEdit: () => void
}

export function ResultsView({ answers, onRestart, onEdit }: ResultsViewProps) {
  const risk = calculateRiskScore(answers)
  const recommendations = generateRecommendations(answers, risk)
  const clientName = (answers.fullName as string) || 'Client'
  const [exporting, setExporting] = useState<string | null>(null)

  async function handleExportPdf() {
    setExporting('pdf')
    try {
      await exportPdf('results-content', clientName)
    } finally {
      setExporting(null)
    }
  }

  async function handleExportDocx() {
    setExporting('docx')
    try {
      await exportDocx(answers)
    } finally {
      setExporting(null)
    }
  }

  return (
    <div className="min-h-screen bg-cream">
      {/* Export buttons - fixed top bar */}
      <div className="bg-white border-b border-cream-dark sticky top-0 z-10 no-print">
        <div className="max-w-3xl mx-auto px-6 py-3 flex items-center justify-between">
          <div className="flex items-center gap-4">
            <button
              onClick={onRestart}
              className="text-sm text-navy-light hover:text-navy-dark transition-colors cursor-pointer"
            >
              &larr; Start Over
            </button>
            <button
              onClick={onEdit}
              className="text-sm text-gold hover:text-gold-light transition-colors cursor-pointer"
            >
              Edit Answers
            </button>
          </div>
          <div className="flex gap-2">
            <button
              onClick={handleExportPdf}
              disabled={exporting !== null}
              className="inline-flex items-center gap-2 bg-navy hover:bg-navy-dark text-white text-sm font-medium py-2 px-4 rounded-lg transition-colors cursor-pointer disabled:opacity-50"
            >
              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
              </svg>
              {exporting === 'pdf' ? 'Generating...' : 'Download PDF'}
            </button>
            <button
              onClick={handleExportDocx}
              disabled={exporting !== null}
              className="inline-flex items-center gap-2 bg-white hover:bg-cream-dark text-navy-dark text-sm font-medium py-2 px-4 rounded-lg border border-gray-200 transition-colors cursor-pointer disabled:opacity-50"
            >
              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
              </svg>
              {exporting === 'docx' ? 'Generating...' : 'Download DOCX'}
            </button>
          </div>
        </div>
      </div>

      {/* Printable content area */}
      <div id="results-content">
        {/* Header */}
        <div className="bg-navy text-white py-8 px-6">
          <div className="max-w-3xl mx-auto">
            <div className="flex items-center justify-between">
              <div>
                <h1 className="text-2xl font-bold mb-1">Client Profile Summary</h1>
                <p className="text-white/60 text-sm">{clientName} &middot; {new Date().toLocaleDateString('en-CA', { year: 'numeric', month: 'long', day: 'numeric' })}</p>
              </div>
              <div className="text-right">
                <div className="text-xs text-white/40 mb-1">The Andrews Group</div>
                <div className="w-12 h-12 bg-white/10 rounded-full flex items-center justify-center">
                  <span className="text-gold font-bold">AG</span>
                </div>
              </div>
            </div>
          </div>
        </div>

        <div className="max-w-3xl mx-auto px-6 py-8">
          {/* Risk Profile */}
          <div className="bg-white rounded-xl shadow-sm p-6 mb-6">
            <h2 className="text-lg font-bold text-navy-dark mb-4">Risk Profile</h2>
            <div className="flex items-center gap-6 mb-4">
              <div
                className="w-24 h-24 rounded-full flex items-center justify-center flex-shrink-0"
                style={{ backgroundColor: `${risk.color}15`, border: `3px solid ${risk.color}` }}
              >
                <span className="text-2xl font-bold" style={{ color: risk.color }}>{risk.score}</span>
              </div>
              <div>
                <h3 className="text-xl font-semibold text-navy-dark mb-1">{risk.category}</h3>
                <p className="text-sm text-gray-500 leading-relaxed">{risk.description}</p>
              </div>
            </div>
            {/* Score bar */}
            <div className="mt-4">
              <div className="flex justify-between text-xs text-gray-400 mb-1">
                <span>Conservative</span>
                <span>Aggressive</span>
              </div>
              <div className="w-full bg-gray-100 rounded-full h-3 relative">
                <div
                  className="h-3 rounded-full transition-all"
                  style={{ width: `${risk.score}%`, backgroundColor: risk.color }}
                />
              </div>
            </div>
          </div>

          {/* Recommendations */}
          {recommendations.length > 0 && (
            <div className="bg-white rounded-xl shadow-sm p-6 mb-6">
              <h2 className="text-lg font-bold text-navy-dark mb-4">Recommendations</h2>
              <div className="space-y-3">
                {recommendations.map((rec, i) => (
                  <RecommendationCard key={i} rec={rec} />
                ))}
              </div>
            </div>
          )}

          {/* Summary Data Sheet */}
          <div className="bg-white rounded-xl shadow-sm p-6 mb-6">
            <h2 className="text-lg font-bold text-navy-dark mb-4">Complete Profile</h2>
            {sections.map(section => (
              <SectionSummary key={section.id} section={section} answers={answers} />
            ))}
          </div>
        </div>
      </div>
    </div>
  )
}

function RecommendationCard({ rec }: { rec: Recommendation }) {
  const colors = {
    high: { bg: 'bg-red-50', border: 'border-red-200', badge: 'bg-red-100 text-red-700' },
    medium: { bg: 'bg-amber-50', border: 'border-amber-200', badge: 'bg-amber-100 text-amber-700' },
    low: { bg: 'bg-blue-50', border: 'border-blue-200', badge: 'bg-blue-100 text-blue-700' },
  }
  const c = colors[rec.priority]

  return (
    <div className={`${c.bg} border ${c.border} rounded-lg p-4`}>
      <div className="flex items-start gap-3">
        <span className={`${c.badge} text-xs font-medium px-2 py-0.5 rounded-full flex-shrink-0 mt-0.5`}>
          {rec.priority}
        </span>
        <div>
          <h4 className="text-sm font-semibold text-navy-dark mb-1">{rec.title}</h4>
          <p className="text-xs text-gray-600 leading-relaxed">{rec.description}</p>
        </div>
      </div>
    </div>
  )
}

function SectionSummary({ section, answers }: { section: typeof sections[0]; answers: Record<string, string | string[]> }) {
  const hasAnswers = section.questions.some(q => {
    const val = answers[q.id]
    return val && (typeof val === 'string' ? val.trim() : val.length > 0)
  })

  if (!hasAnswers) return null

  return (
    <div className="mb-6 last:mb-0">
      <h3 className="text-sm font-semibold text-gold uppercase tracking-wide mb-3 pb-2 border-b border-cream-dark">
        {section.title}
      </h3>
      <div className="space-y-2">
        {section.questions.map(question => {
          const val = answers[question.id]
          const noteVal = answers[`${question.id}_notes`] as string

          if (!val || (typeof val === 'string' && !val.trim()) || (Array.isArray(val) && val.length === 0)) {
            return null
          }

          let displayValue: string
          if (Array.isArray(val)) {
            displayValue = val.map(v => {
              const opt = question.options?.find(o => o.value === v)
              return opt?.label ?? v
            }).join(', ')
          } else if (question.options) {
            const opt = question.options.find(o => o.value === val)
            displayValue = opt?.label ?? val
          } else if (question.type === 'scale') {
            displayValue = `${val} / ${question.max}`
          } else {
            displayValue = val
          }

          return (
            <div key={question.id} className="flex gap-2 text-sm">
              <span className="text-gray-400 min-w-[180px] flex-shrink-0">{question.label}</span>
              <span className="text-navy-dark font-medium">
                {displayValue}
                {noteVal && <span className="text-gray-400 font-normal ml-2">({noteVal})</span>}
              </span>
            </div>
          )
        })}
      </div>
    </div>
  )
}
