import { useState } from 'react'
import './App.css'
import {
  buildGoogleSheetCsvUrl,
  defaultSheetSource,
  formatHintBookFromAppsScript,
  formatHintBookFromCsv,
  sampleAppsScriptResponse,
  sampleCsv,
  sheetColumnGuide,
  type FormatResult,
  type FormattedPage,
  type RenderSettings,
  type RichTextRun,
} from './lib/hintbook'

const defaultAppsScriptSource =
  'https://script.google.com/macros/s/AKfycbxujRiMhpqRckPyMcnWahn4eJ6cvl3SVG5rGGao7_55iWpjKq5duNBHIeVtH4MyzifmFw/exec'

const BASE_FONT_SIZE = 10
const STEP_FONT_SCALE = 4
const BODY_FONT_SCALE = 2.5

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

async function fetchAppsScriptJson(source: string) {
  const response = await fetch(source)

  if (!response.ok) {
    throw new Error(`Failed to fetch Apps Script JSON: ${response.status}`)
  }

  return response.json()
}

function extractGoogleDriveFileId(source: string) {
  const trimmed = source.trim()
  if (!trimmed) {
    return null
  }

  if (/^[a-zA-Z0-9_-]{20,}$/.test(trimmed) && !trimmed.includes('/')) {
    return trimmed
  }

  try {
    const url = new URL(trimmed)
    const host = url.hostname.replace(/^www\./, '')

    if (host === 'drive.google.com') {
      const id = url.searchParams.get('id')
      if (id) {
        return id
      }

      const fileMatch = url.pathname.match(/\/file\/d\/([a-zA-Z0-9_-]+)/)
      if (fileMatch) {
        return fileMatch[1]
      }
    }
  } catch {
    return null
  }

  return null
}

function resolveImageSource(source: string) {
  const trimmed = source.trim()
  if (!trimmed) {
    return null
  }

  const driveFileId = extractGoogleDriveFileId(trimmed)
  if (driveFileId) {
    return `https://drive.google.com/thumbnail?id=${driveFileId}&sz=w1600`
  }

  try {
    return new URL(trimmed).toString()
  } catch {
    return null
  }
}

function scaledFontSize(size: number | undefined, multiplier: number) {
  const baseSize = size ?? BASE_FONT_SIZE
  return `${baseSize * multiplier}px`
}

function resolveRunFontFamily(
  run: RichTextRun,
  defaultFontFamily?: string,
  cellFontFamily?: string,
) {
  return (
    run.fontFamily?.trim() ??
    cellFontFamily?.trim() ??
    defaultFontFamily?.trim()
  )
}

function resolveTextDecoration(
  underline?: boolean,
  strikethrough?: boolean,
) {
  const values = [
    underline ? 'underline' : '',
    strikethrough ? 'line-through' : '',
  ].filter(Boolean)

  if (values.length > 0) {
    return values.join(' ')
  }

  if (underline !== undefined || strikethrough !== undefined) {
    return 'none'
  }

  return undefined
}

function runStyle(
  run: RichTextRun,
  multiplier: number,
  options: {
    defaultFontFamily?: string
    cellFontFamily?: string
    baseFontSize?: number
    baseTextColor?: string
    omitFontFamily?: boolean
  },
) {
  return {
    color: run.textColor ?? options.baseTextColor,
    fontFamily: options.omitFontFamily
      ? undefined
      : resolveRunFontFamily(run, options.defaultFontFamily, options.cellFontFamily),
    fontSize: scaledFontSize(run.fontSize ?? options.baseFontSize, multiplier),
    fontWeight: run.bold === undefined ? undefined : run.bold ? '700' : '400',
    fontStyle: run.italic === undefined ? undefined : run.italic ? 'italic' : 'normal',
    textDecoration: resolveTextDecoration(run.underline, run.strikethrough),
  }
}

function RichText({
  runs,
  fallback,
  className,
  multiplier,
  defaultFontFamily,
  cellFontFamily,
  baseFontSize,
  baseTextColor,
  omitFontFamily,
}: {
  runs?: RichTextRun[]
  fallback: string
  className: string
  multiplier: number
  defaultFontFamily?: string
  cellFontFamily?: string
  baseFontSize?: number
  baseTextColor?: string
  omitFontFamily?: boolean
}) {
  const sourceRuns =
    runs && runs.length > 0
      ? runs
      : [
          {
            text: fallback,
          },
        ]

  return (
    <div className={className}>
      {sourceRuns.map((run, runIndex) => {
        const parts = run.text.split('\n')
        return (
          <span
            key={`${runIndex}-${run.text}`}
            style={runStyle(run, multiplier, {
              defaultFontFamily,
              cellFontFamily,
              baseFontSize,
              baseTextColor,
              omitFontFamily,
            })}
          >
            {parts.map((part, partIndex) => (
              <span key={`${runIndex}-${partIndex}`}>
                {part}
                {partIndex < parts.length - 1 ? <br /> : null}
              </span>
            ))}
          </span>
        )
      })}
    </div>
  )
}

function PageImage({ source }: { source: string }) {
  const [failed, setFailed] = useState(false)
  const resolvedSource = resolveImageSource(source)

  if (!resolvedSource || failed) {
    return null
  }

  return (
    <div className="pageImageFrame">
      <img
        className="pageImage"
        src={resolvedSource}
        alt=""
        loading="lazy"
        onError={() => setFailed(true)}
      />
    </div>
  )
}

function PagePreview({
  page,
  position,
  settings,
}: {
  page: FormattedPage | null
  position: 'left' | 'right'
  settings?: RenderSettings
}) {
  const stepStyle = page?.stepStyle
  const bodyStyle = page?.bodyStyle
  const stepMultiplier = settings?.stepFontScale ?? STEP_FONT_SCALE
  const bodyMultiplier = settings?.bodyFontScale ?? BODY_FONT_SCALE
  const stepFontFamily = settings?.stepFontFamily ?? stepStyle?.fontFamily
  const bodyFontFamily = settings?.bodyFontFamily ?? bodyStyle?.fontFamily
  const sideFontFamily = settings?.sideFontFamily
  const pageNoFontFamily = settings?.pageNoFontFamily
  const stepBlockStyle = {
    backgroundColor: stepStyle?.backgroundColor,
    color: stepStyle?.textColor,
    fontFamily: stepFontFamily,
    fontSize: scaledFontSize(stepStyle?.fontSize, stepMultiplier),
    fontWeight: stepStyle?.bold ? '700' : undefined,
    fontStyle: stepStyle?.italic ? 'italic' : undefined,
    textDecoration: [
      stepStyle?.underline ? 'underline' : '',
      stepStyle?.strikethrough ? 'line-through' : '',
    ]
      .filter(Boolean)
      .join(' ') || undefined,
  }
  const bodyBlockStyle = {
    borderColor: stepStyle?.backgroundColor,
    backgroundColor: bodyStyle?.backgroundColor,
    color: bodyStyle?.textColor,
    fontFamily: bodyFontFamily,
    fontSize: scaledFontSize(bodyStyle?.fontSize, bodyMultiplier),
    fontWeight: bodyStyle?.bold ? '700' : undefined,
    fontStyle: bodyStyle?.italic ? 'italic' : undefined,
    textDecoration: [
      bodyStyle?.underline ? 'underline' : '',
      bodyStyle?.strikethrough ? 'line-through' : '',
    ]
      .filter(Boolean)
      .join(' ') || undefined,
  }
  const sideLabelStyle = {
    fontFamily: sideFontFamily,
  }
  const pageNumberStyle = {
    fontFamily: pageNoFontFamily,
  }

  if (!page) {
    return (
      <article className={`pagePreview pagePreview-${position}`}>
        <div className="pageStep" />
        <div className={`pageSideLabel pageSideLabel-${position}`} />
        <div className={`pageBody pageBody-${position}`} />
        <div className="pageNumber" />
      </article>
    )
  }

  return (
    <article className={`pagePreview pagePreview-${position}`}>
      <div className="pageStep" style={stepBlockStyle}>
        <RichText
          runs={page.stepRuns}
          fallback={page.step || ''}
          className="pageStepText"
          multiplier={stepMultiplier}
          defaultFontFamily={settings?.stepFontFamily}
          cellFontFamily={stepStyle?.fontFamily}
          baseFontSize={stepStyle?.fontSize}
          baseTextColor={stepStyle?.textColor}
          omitFontFamily
        />
      </div>
      <div
        className={`pageSideLabel pageSideLabel-${position}`}
        style={sideLabelStyle}
      >
        {page.side || ''}
      </div>
      <div className={`pageBody pageBody-${position}`} style={bodyBlockStyle}>
        <RichText
          runs={page.bodyRuns}
          fallback={page.body || ''}
          className="pageBodyText"
          multiplier={bodyMultiplier}
          defaultFontFamily={settings?.bodyFontFamily}
          cellFontFamily={bodyStyle?.fontFamily}
          baseFontSize={bodyStyle?.fontSize}
          baseTextColor={bodyStyle?.textColor}
        />
        {page.image ? (
          <PageImage source={page.image} />
        ) : null}
      </div>
      <div className="pageNumber" style={pageNumberStyle}>
        {page.pageNo || ''}
      </div>
    </article>
  )
}

function App() {
  const [sheetSource, setSheetSource] = useState(defaultSheetSource)
  const [appsScriptSource, setAppsScriptSource] = useState(defaultAppsScriptSource)
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

  const handleLoadAppsScript = async () => {
    setLoading(true)
    try {
      const payload = await fetchAppsScriptJson(appsScriptSource)
      const next = formatHintBookFromAppsScript(payload)
      setResult(next)
      setStatus(
        `Loaded Apps Script data and formatted ${next.pages.length} pages into ${next.spreads.length} print spreads.`,
      )
    } catch (error) {
      setStatus(error instanceof Error ? error.message : 'Apps Script load failed.')
    } finally {
      setLoading(false)
    }
  }

  const handleUseAppsScriptSample = () => {
    const next = formatHintBookFromAppsScript(sampleAppsScriptResponse)
    setResult(next)
    setStatus(
      `Loaded sample Apps Script data and formatted ${next.pages.length} pages into ${next.spreads.length} print spreads.`,
    )
  }

  const handleExportPdf = () => {
    if (!result?.spreads.length) {
      setStatus('No print spreads available for PDF export.')
      return
    }

    setStatus(
      `Opening print dialog for ${result.spreads.length} sheets. Choose "Save as PDF" to download.`,
    )
    window.print()
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
          <button
            type="button"
            onClick={handleExportPdf}
            disabled={!result?.spreads.length}
          >
            Export PDF
          </button>
        </div>

        <label>
          Apps Script Web App URL
          <input
            type="text"
            value={appsScriptSource}
            onChange={(event) => setAppsScriptSource(event.target.value)}
            placeholder="https://script.google.com/macros/s/.../exec"
          />
        </label>

        <div className="buttonRow">
          <button type="button" onClick={handleLoadAppsScript} disabled={loading}>
            {loading ? 'Loading...' : 'Load Apps Script'}
          </button>
          <button type="button" onClick={handleUseAppsScriptSample}>
            Use Apps Script sample
          </button>
        </div>

        <p className="note">
          CSV columns: <code>order, page_no, step, side, body, image</code>
        </p>
        <p className="note">
          Apps Script route can also read the step cell background color, text color,
          and font family.
        </p>
        <p className="note">
          <code>image</code> accepts Google Drive share links and regular image URLs.
        </p>
        <p className="note">
          PDF export uses the browser print dialog. Select <code>Save as PDF</code>.
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
            <div className="sheetViewport">
              <div className="sheetCanvas">
                <PagePreview
                  page={spread.leftPage}
                  position="left"
                  settings={result?.settings}
                />
                <PagePreview
                  page={spread.rightPage}
                  position="right"
                  settings={result?.settings}
                />
              </div>
            </div>
          </section>
        )) ?? null}
      </section>
    </main>
  )
}

export default App
