import { useState } from 'react'
import './App.css'
import {
  buildGoogleSheetCsvUrl,
  defaultSheetSource,
  formatHintBookFromCsv,
  sampleCsv,
  sheetColumnGuide,
  type FormatResult,
  type FormattedPage,
} from './lib/hintbook'

function downloadJson(result: FormatResult) {
  const blob = new Blob([JSON.stringify(result, null, 2)], {
    type: 'application/json',
  })
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = 'hintbook-formatted.json'
  anchor.click()
  URL.revokeObjectURL(url)
}

async function fetchSpreadsheetCsv(source: string) {
  const csvUrl = buildGoogleSheetCsvUrl(source)
  const response = await fetch(csvUrl)

  if (!response.ok) {
    throw new Error(`Failed to fetch CSV: ${response.status}`)
  }

  return response.text()
}

function PagePreview({
  page,
  position,
}: {
  page: FormattedPage | null
  position: 'left' | 'right'
}) {
  if (!page) {
    return (
      <article className={`pagePreview pagePreview-${position}`}>
        <div className="pageStep">empty</div>
        <div className="pageContent">
          <div className="pageSideLabel">side</div>
          <div className="pageBody">No page assigned.</div>
        </div>
        <div className="pageNumber">-</div>
      </article>
    )
  }

  return (
    <article className={`pagePreview pagePreview-${position}`}>
      <div className="pageStep">{page.step || 'step'}</div>
      <div className="pageContent">
        <div className="pageSideLabel">{page.side || 'side'}</div>
        <div className="pageBody">
          <pre>{page.body || 'body'}</pre>
          {page.image ? (
            <div className="pageImagePlaceholder">image: {page.image}</div>
          ) : null}
        </div>
      </div>
      <div className="pageNumber">{page.pageNo || '-'}</div>
    </article>
  )
}

function App() {
  const [sheetSource, setSheetSource] = useState(defaultSheetSource)
  const [csvText, setCsvText] = useState(sampleCsv)
  const [result, setResult] = useState<FormatResult | null>(() =>
    formatHintBookFromCsv(sampleCsv),
  )
  const [status, setStatus] = useState('Ready.')
  const [loading, setLoading] = useState(false)

  const handleFormat = () => {
    try {
      const next = formatHintBookFromCsv(csvText)
      setResult(next)
      setStatus(
        `Formatted ${next.pages.length} pages into ${next.spreads.length} print spreads.`,
      )
    } catch (error) {
      setResult(null)
      setStatus(error instanceof Error ? error.message : 'Formatting failed.')
    }
  }

  const handleLoadSpreadsheet = async () => {
    setLoading(true)
    try {
      const fetchedCsv = await fetchSpreadsheetCsv(sheetSource)
      setCsvText(fetchedCsv)
      const next = formatHintBookFromCsv(fetchedCsv)
      setResult(next)
      setStatus(
        `Loaded spreadsheet and formatted ${next.pages.length} pages into ${next.spreads.length} print spreads.`,
      )
    } catch (error) {
      setStatus(error instanceof Error ? error.message : 'Spreadsheet load failed.')
    } finally {
      setLoading(false)
    }
  }

  return (
    <main className="app">
      <section className="controls">
        <h1>Hint Book Formatter</h1>
        <p>Spreadsheet input first. Temporary layout only.</p>

        <label>
          Google Sheets URL or ID
          <input
            type="text"
            value={sheetSource}
            onChange={(event) => setSheetSource(event.target.value)}
          />
        </label>

        <div className="buttonRow">
          <button type="button" onClick={handleLoadSpreadsheet} disabled={loading}>
            {loading ? 'Loading...' : 'Load spreadsheet'}
          </button>
          <button type="button" onClick={() => setCsvText(sampleCsv)}>
            Use sample CSV
          </button>
          <button
            type="button"
            onClick={() => result && downloadJson(result)}
            disabled={!result}
          >
            Download JSON
          </button>
        </div>

        <p className="note">
          Expected columns: <code>order, page_no, step, side, body, image</code>
        </p>

        <label>
          CSV text
          <textarea
            value={csvText}
            onChange={(event) => setCsvText(event.target.value)}
            rows={12}
          />
        </label>

        <div className="buttonRow">
          <button type="button" onClick={handleFormat}>
            Format CSV
          </button>
        </div>

        <h2>Status</h2>
        <pre>{status}</pre>

        <h2>Columns</h2>
        <table>
          <thead>
            <tr>
              <th>column</th>
              <th>meaning</th>
            </tr>
          </thead>
          <tbody>
            {sheetColumnGuide.map(([column, description]) => (
              <tr key={column}>
                <td>{column}</td>
                <td>{description}</td>
              </tr>
            ))}
          </tbody>
        </table>

        {result?.warnings.length ? (
          <>
            <h2>Warnings</h2>
            <ul>
              {result.warnings.map((warning) => (
                <li key={warning}>{warning}</li>
              ))}
            </ul>
          </>
        ) : null}
      </section>

      <section className="previewArea">
        <h2>Print spreads</h2>
        <p className="note">
          Order mapping: 1-4 becomes [1,4] then [3,2]. The same rule repeats for
          5-8, 9-12...
        </p>

        {result?.spreads.map((spread) => (
          <section key={spread.sheetNumber} className="sheetPreview">
            <div className="sheetMeta">
              <span>sheet {spread.sheetNumber}</span>
              <span>spread group {spread.spreadNumber}</span>
            </div>
            <div className="sheetCanvas">
              <PagePreview page={spread.leftPage} position="left" />
              <PagePreview page={spread.rightPage} position="right" />
            </div>
          </section>
        )) ?? null}
      </section>
    </main>
  )
}

export default App
