import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  BorderStyle,
  TableCell,
  TableRow,
  Table,
  WidthType,
} from 'docx'
import { saveAs } from 'file-saver'
import { sections } from './questions'
import { calculateRiskScore } from './scoring'
import { generateRecommendations } from './recommendations'

export async function exportDocx(answers: Record<string, string | string[]>) {
  const risk = calculateRiskScore(answers)
  const recommendations = generateRecommendations(answers, risk)
  const clientName = (answers.fullName as string) || 'Client'
  const date = new Date().toLocaleDateString('en-CA', { year: 'numeric', month: 'long', day: 'numeric' })

  const docSections: (Paragraph | Table)[] = []

  // Header
  docSections.push(
    new Paragraph({
      children: [new TextRun({ text: 'The Andrews Group', bold: true, size: 28, color: '1e3a5f' })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
    }),
    new Paragraph({
      children: [new TextRun({ text: 'CI Assante Wealth Management', size: 18, color: '888888' })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
    }),
    new Paragraph({
      children: [new TextRun({ text: 'Client Profile Summary', bold: true, size: 32, color: '1e3a5f' })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: `${clientName}  •  ${date}`, size: 20, color: '888888' })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: 'dddddd' } },
    })
  )

  // Risk Profile
  docSections.push(
    new Paragraph({
      text: 'Risk Profile',
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 300, after: 100 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: `Score: ${risk.score}/100`, bold: true, size: 24 }),
        new TextRun({ text: `  —  ${risk.category}`, size: 24, color: '1e3a5f' }),
      ],
      spacing: { after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: risk.description, size: 20, color: '555555' })],
      spacing: { after: 300 },
    })
  )

  // Recommendations
  if (recommendations.length > 0) {
    docSections.push(
      new Paragraph({
        text: 'Recommendations',
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 300, after: 100 },
      })
    )

    const recRows = recommendations.map(
      rec =>
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: rec.priority.toUpperCase(), bold: true, size: 18 })] })],
              width: { size: 15, type: WidthType.PERCENTAGE },
            }),
            new TableCell({
              children: [
                new Paragraph({ children: [new TextRun({ text: rec.title, bold: true, size: 20 })], spacing: { after: 40 } }),
                new Paragraph({ children: [new TextRun({ text: rec.description, size: 18, color: '555555' })] }),
              ],
              width: { size: 85, type: WidthType.PERCENTAGE },
            }),
          ],
        })
    )

    docSections.push(
      new Table({ rows: recRows, width: { size: 100, type: WidthType.PERCENTAGE } }),
      new Paragraph({ text: '', spacing: { after: 300 } })
    )
  }

  // Complete Profile sections
  for (const section of sections) {
    const answeredQuestions = section.questions.filter(q => {
      const val = answers[q.id]
      return val && (typeof val === 'string' ? val.trim() : val.length > 0)
    })

    if (answeredQuestions.length === 0) continue

    docSections.push(
      new Paragraph({
        text: section.title,
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 300, after: 100 },
      })
    )

    const rows = answeredQuestions.map(question => {
      const val = answers[question.id]
      const noteVal = answers[`${question.id}_notes`] as string

      let displayValue: string
      if (Array.isArray(val)) {
        displayValue = val.map(v => {
          const opt = question.options?.find(o => o.value === v)
          return opt?.label ?? v
        }).join(', ')
      } else if (question.options) {
        const opt = question.options.find(o => o.value === val)
        displayValue = opt?.label ?? (val as string)
      } else if (question.type === 'scale') {
        displayValue = `${val} / ${question.max}`
      } else {
        displayValue = val as string
      }

      if (noteVal) displayValue += ` (${noteVal})`

      return new TableRow({
        children: [
          new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: question.label, size: 18, color: '888888' })] })],
            width: { size: 40, type: WidthType.PERCENTAGE },
          }),
          new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: displayValue, size: 18, bold: true })] })],
            width: { size: 60, type: WidthType.PERCENTAGE },
          }),
        ],
      })
    })

    docSections.push(new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }))
  }

  const doc = new Document({
    sections: [{ children: docSections }],
  })

  const blob = await Packer.toBlob(doc)
  const fileDate = new Date().toISOString().split('T')[0]
  saveAs(blob, `${clientName.replace(/\s+/g, '_')}_Profile_${fileDate}.docx`)
}
