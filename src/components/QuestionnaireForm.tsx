import { useState } from 'react'
import { sections, type Question } from '../lib/questions'

interface QuestionnaireFormProps {
  initialAnswers?: Record<string, string | string[]>
  onComplete: (answers: Record<string, string | string[]>) => void
}

export function QuestionnaireForm({ initialAnswers, onComplete }: QuestionnaireFormProps) {
  const [currentStep, setCurrentStep] = useState(0)
  const [answers, setAnswers] = useState<Record<string, string | string[]>>(
    () => {
      if (!initialAnswers) return {}
      // Separate notes from answers
      const ans: Record<string, string | string[]> = {}
      for (const [key, val] of Object.entries(initialAnswers)) {
        if (!key.endsWith('_notes')) ans[key] = val
      }
      return ans
    }
  )
  const [notes, setNotes] = useState<Record<string, string>>(
    () => {
      if (!initialAnswers) return {}
      const n: Record<string, string> = {}
      for (const [key, val] of Object.entries(initialAnswers)) {
        if (key.endsWith('_notes')) n[key.replace('_notes', '')] = val as string
      }
      return n
    }
  )

  const section = sections[currentStep]
  const isLastStep = currentStep === sections.length - 1

  function updateAnswer(questionId: string, value: string | string[]) {
    setAnswers(prev => ({ ...prev, [questionId]: value }))
  }

  function updateNote(questionId: string, value: string) {
    setNotes(prev => ({ ...prev, [questionId]: value }))
  }

  function toggleMultiselect(questionId: string, value: string) {
    const current = (answers[questionId] as string[]) || []
    if (value === 'none') {
      updateAnswer(questionId, ['none'])
      return
    }
    const filtered = current.filter(v => v !== 'none')
    if (filtered.includes(value)) {
      updateAnswer(questionId, filtered.filter(v => v !== value))
    } else {
      updateAnswer(questionId, [...filtered, value])
    }
  }

  function handleNext() {
    if (isLastStep) {
      // Merge notes into answers with _notes suffix
      const finalAnswers = { ...answers }
      for (const [key, value] of Object.entries(notes)) {
        if (value.trim()) {
          finalAnswers[`${key}_notes`] = value
        }
      }
      onComplete(finalAnswers)
    } else {
      setCurrentStep(prev => prev + 1)
      window.scrollTo(0, 0)
    }
  }

  function handleBack() {
    setCurrentStep(prev => prev - 1)
    window.scrollTo(0, 0)
  }

  return (
    <div className="min-h-screen bg-cream">
      {/* Progress bar */}
      <div className="bg-white border-b border-cream-dark sticky top-0 z-10 no-print">
        <div className="max-w-2xl mx-auto px-6 py-4">
          <div className="flex items-center justify-between mb-2">
            <span className="text-sm font-medium text-navy-dark">{section.title}</span>
            <span className="text-xs text-gray-400">{currentStep + 1} of {sections.length}</span>
          </div>
          <div className="w-full bg-cream-dark rounded-full h-2">
            <div
              className="bg-gold h-2 rounded-full transition-all duration-300"
              style={{ width: `${((currentStep + 1) / sections.length) * 100}%` }}
            />
          </div>
        </div>
      </div>

      {/* Questions */}
      <div className="max-w-2xl mx-auto px-6 py-8">
        <div className="mb-8">
          <h2 className="text-2xl font-bold text-navy-dark mb-1">{section.title}</h2>
          <p className="text-gray-500 text-sm">{section.description}</p>
        </div>

        <div className="space-y-6">
          {section.questions.map(question => (
            <QuestionField
              key={question.id}
              question={question}
              value={answers[question.id]}
              note={notes[question.id]}
              onChange={value => updateAnswer(question.id, value)}
              onToggle={value => toggleMultiselect(question.id, value)}
              onNoteChange={value => updateNote(question.id, value)}
            />
          ))}
        </div>

        {/* Navigation */}
        <div className="flex justify-between mt-10 no-print">
          <button
            onClick={handleBack}
            disabled={currentStep === 0}
            className="px-5 py-2.5 text-sm font-medium text-navy-light hover:text-navy-dark disabled:opacity-30 disabled:cursor-not-allowed transition-colors cursor-pointer"
          >
            Back
          </button>
          <button
            onClick={handleNext}
            className="bg-navy hover:bg-navy-dark text-white font-medium py-2.5 px-8 rounded-lg transition-colors cursor-pointer"
          >
            {isLastStep ? 'View Results' : 'Next'}
          </button>
        </div>
      </div>
    </div>
  )
}

interface QuestionFieldProps {
  question: Question
  value: string | string[] | undefined
  note: string | undefined
  onChange: (value: string) => void
  onToggle: (value: string) => void
  onNoteChange: (value: string) => void
}

function QuestionField({ question, value, note, onChange, onToggle, onNoteChange }: QuestionFieldProps) {
  const baseInputClass = 'w-full border border-gray-200 rounded-lg px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-gold/50 focus:border-gold bg-white'

  return (
    <div className="bg-white rounded-xl p-5 shadow-sm">
      <label className="block text-sm font-medium text-navy-dark mb-2">
        {question.label}
        {question.required && <span className="text-red-400 ml-0.5">*</span>}
      </label>

      {question.type === 'text' && (
        <input
          type="text"
          className={baseInputClass}
          placeholder={question.placeholder}
          value={(value as string) || ''}
          onChange={e => onChange(e.target.value)}
        />
      )}

      {question.type === 'number' && (
        <input
          type="number"
          className={baseInputClass}
          placeholder={question.placeholder}
          min={question.min}
          max={question.max}
          value={(value as string) || ''}
          onChange={e => onChange(e.target.value)}
        />
      )}

      {question.type === 'textarea' && (
        <textarea
          className={`${baseInputClass} min-h-[80px] resize-y`}
          placeholder={question.placeholder}
          value={(value as string) || ''}
          onChange={e => onChange(e.target.value)}
        />
      )}

      {question.type === 'select' && question.options && (
        <div className="space-y-1.5">
          {question.options.map(option => (
            <label
              key={option.value}
              className={`flex items-center gap-3 px-4 py-2.5 rounded-lg border cursor-pointer transition-colors text-sm ${
                value === option.value
                  ? 'border-gold bg-gold/5 text-navy-dark'
                  : 'border-gray-100 hover:border-gray-200 text-gray-600'
              }`}
            >
              <input
                type="radio"
                name={question.id}
                value={option.value}
                checked={value === option.value}
                onChange={() => onChange(option.value)}
                className="sr-only"
              />
              <div className={`w-4 h-4 rounded-full border-2 flex items-center justify-center flex-shrink-0 ${
                value === option.value ? 'border-gold' : 'border-gray-300'
              }`}>
                {value === option.value && <div className="w-2 h-2 rounded-full bg-gold" />}
              </div>
              {option.label}
            </label>
          ))}
        </div>
      )}

      {question.type === 'multiselect' && question.options && (
        <div className="space-y-1.5">
          {question.options.map(option => {
            const selected = ((value as string[]) || []).includes(option.value)
            return (
              <label
                key={option.value}
                className={`flex items-center gap-3 px-4 py-2.5 rounded-lg border cursor-pointer transition-colors text-sm ${
                  selected
                    ? 'border-gold bg-gold/5 text-navy-dark'
                    : 'border-gray-100 hover:border-gray-200 text-gray-600'
                }`}
              >
                <input
                  type="checkbox"
                  checked={selected}
                  onChange={() => onToggle(option.value)}
                  className="sr-only"
                />
                <div className={`w-4 h-4 rounded border-2 flex items-center justify-center flex-shrink-0 ${
                  selected ? 'border-gold bg-gold' : 'border-gray-300'
                }`}>
                  {selected && (
                    <svg className="w-2.5 h-2.5 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={3} d="M5 13l4 4L19 7" />
                    </svg>
                  )}
                </div>
                {option.label}
              </label>
            )
          })}
        </div>
      )}

      {question.type === 'scale' && (
        <div>
          <div className="flex items-center gap-4">
            <input
              type="range"
              min={question.min}
              max={question.max}
              step={question.step}
              value={(value as string) || String(question.min)}
              onChange={e => onChange(e.target.value)}
              className="flex-1 accent-gold"
            />
            <span className="text-2xl font-bold text-navy-dark w-10 text-center">
              {value || question.min}
            </span>
          </div>
          {question.scaleLabels && (
            <div className="flex justify-between text-xs text-gray-400 mt-1 px-1">
              <span>{question.scaleLabels.low}</span>
              <span>{question.scaleLabels.high}</span>
            </div>
          )}
        </div>
      )}

      {question.hasNotes && (
        <div className="mt-3 pt-3 border-t border-gray-100">
          <label className="text-xs text-gray-400 mb-1 block">Additional notes (optional)</label>
          <input
            type="text"
            className="w-full border border-gray-100 rounded-lg px-3 py-2 text-xs focus:outline-none focus:ring-1 focus:ring-gold/30 bg-cream/50"
            placeholder="Any additional context..."
            value={note || ''}
            onChange={e => onNoteChange(e.target.value)}
          />
        </div>
      )}
    </div>
  )
}
