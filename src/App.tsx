import { useState } from 'react'
import { Welcome } from './components/Welcome'
import { QuestionnaireForm } from './components/QuestionnaireForm'
import { ResultsView } from './components/ResultsView'

type Screen = 'welcome' | 'form' | 'results'

function App() {
  const [screen, setScreen] = useState<Screen>('welcome')
  const [answers, setAnswers] = useState<Record<string, string | string[]>>({})

  function handleComplete(data: Record<string, string | string[]>) {
    setAnswers(data)
    setScreen('results')
    window.scrollTo(0, 0)
  }

  function handleRestart() {
    setAnswers({})
    setScreen('welcome')
    window.scrollTo(0, 0)
  }

  function handleEdit() {
    setScreen('form')
    window.scrollTo(0, 0)
  }

  if (screen === 'welcome') return <Welcome onStart={() => setScreen('form')} />
  if (screen === 'form') return <QuestionnaireForm initialAnswers={answers} onComplete={handleComplete} />
  return <ResultsView answers={answers} onRestart={handleRestart} onEdit={handleEdit} />
}

export default App
