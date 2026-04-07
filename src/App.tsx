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
  type ImageSources,
  type RenderSettings,
  type RichTextRun,
} from './lib/hintbook'

const defaultAppsScriptSource =
  'https://script.google.com/macros/s/AKfycbxujRiMhpqRckPyMcnWahn4eJ6cvl3SVG5rGGao7_55iWpjKq5duNBHIeVtH4MyzifmFw/exec'

const BASE_FONT_SIZE = 10
const STEP_FONT_SCALE = 4
const BODY_FONT_SCALE = 2.5
const INLINE_IMAGE_TOKEN = /\{\{(images?|img)(?::([0-9,\s]+))?(?:\s+([^}]+))?\}\}/gi

type InlineImageOptions = {
  imageKey?: string
  imageKeys?: string[]
  source?: string
  width?: string
  height?: string
  align?: string
  fit?: string
}

type BodyContentItem =
  | {
      type: 'text'
      runs: RichTextRun[]
    }
  | {
      type: 'image'
      options: InlineImageOptions
    }
  | {
      type: 'imageRow'
      options: InlineImageOptions
    }

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

function resolveImageSource(source?: string) {
  const trimmed = source?.trim()
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

function normalizeCssSize(value?: string) {
  const trimmed = value?.trim()
  if (!trimmed) {
    return undefined
  }

  if (/^-?\d+(\.\d+)?$/.test(trimmed)) {
    return `${trimmed}px`
  }

  return trimmed
}

function resolveImagePosition(value?: string) {
  const normalized = value?.trim().toLowerCase()
  return normalized === 'top' ? 'top' : 'bottom'
}

function resolveImageAlign(value?: string) {
  const normalized = value?.trim().toLowerCase()

  if (normalized === 'left') {
    return 'flex-start'
  }

  if (normalized === 'right') {
    return 'flex-end'
  }

  return 'center'
}

function resolveImageFit(value?: string) {
  const normalized = value?.trim().toLowerCase()
  return normalized === 'cover' ? 'cover' : 'contain'
}

function cloneRunWithText(run: RichTextRun, text: string): RichTextRun {
  return {
    ...run,
    text,
  }
}

function sliceRunsByRange(
  runs: RichTextRun[],
  startIndex: number,
  endIndex: number,
) {
  if (startIndex >= endIndex) {
    return []
  }

  const slicedRuns: RichTextRun[] = []
  let cursor = 0

  for (const run of runs) {
    const nextCursor = cursor + run.text.length
    const overlapStart = Math.max(startIndex, cursor)
    const overlapEnd = Math.min(endIndex, nextCursor)

    if (overlapStart < overlapEnd) {
      const relativeStart = overlapStart - cursor
      const relativeEnd = overlapEnd - cursor
      const text = run.text.slice(relativeStart, relativeEnd)

      if (text) {
        slicedRuns.push(cloneRunWithText(run, text))
      }
    }

    cursor = nextCursor
  }

  return slicedRuns
}

function trimLeadingLineBreak(runs: RichTextRun[]) {
  if (runs.length === 0) {
    return runs
  }

  const nextRuns = [...runs]
  const firstRun = nextRuns[0]
  if (!firstRun) {
    return nextRuns
  }

  let trimmedText = firstRun.text
  if (trimmedText.startsWith('\r\n')) {
    trimmedText = trimmedText.slice(2)
  } else if (trimmedText.startsWith('\n') || trimmedText.startsWith('\r')) {
    trimmedText = trimmedText.slice(1)
  }

  if (trimmedText === firstRun.text) {
    return nextRuns
  }

  if (trimmedText) {
    nextRuns[0] = cloneRunWithText(firstRun, trimmedText)
    return nextRuns
  }

  nextRuns.shift()
  return nextRuns
}

function trimTrailingLineBreak(runs: RichTextRun[]) {
  if (runs.length === 0) {
    return runs
  }

  const nextRuns = [...runs]
  const lastIndex = nextRuns.length - 1
  const lastRun = nextRuns[lastIndex]
  if (!lastRun) {
    return nextRuns
  }

  let trimmedText = lastRun.text
  if (trimmedText.endsWith('\r\n')) {
    trimmedText = trimmedText.slice(0, -2)
  } else if (trimmedText.endsWith('\n') || trimmedText.endsWith('\r')) {
    trimmedText = trimmedText.slice(0, -1)
  }

  if (trimmedText === lastRun.text) {
    return nextRuns
  }

  if (trimmedText) {
    nextRuns[lastIndex] = cloneRunWithText(lastRun, trimmedText)
    return nextRuns
  }

  nextRuns.pop()
  return nextRuns
}

function parseInlineImageOptions(input?: string): InlineImageOptions {
  const options: InlineImageOptions = {}
  if (!input) {
    return options
  }

  const pairs = input.match(/[^\s"]+="[^"]*"|[^\s]+/g) ?? []

  for (const pair of pairs) {
    const separatorIndex = pair.indexOf('=')
    if (separatorIndex <= 0) {
      continue
    }

    const key = pair.slice(0, separatorIndex).trim().toLowerCase()
    const rawValue = pair.slice(separatorIndex + 1).trim().replace(/^"(.*)"$/, '$1')
    if (!rawValue) {
      continue
    }

    if (key === 'src' || key === 'source' || key === 'url') {
      options.source = rawValue
    } else if (key === 'index' || key === 'image' || key === 'key') {
      options.imageKey = rawValue
    } else if (key === 'width' || key === 'w') {
      options.width = rawValue
    } else if (key === 'height' || key === 'h') {
      options.height = rawValue
    } else if (key === 'align') {
      options.align = rawValue
    } else if (key === 'fit') {
      options.fit = rawValue
    }
  }

  return options
}

function parseInlineImageKeys(input?: string) {
  if (!input) {
    return undefined
  }

  const keys = input
    .split(',')
    .map((value) => value.trim())
    .filter(Boolean)

  return keys.length > 0 ? keys : undefined
}

function buildImageSourceMap(page: FormattedPage | null): ImageSources {
  const sources: ImageSources = {
    ...(page?.imageSources ?? {}),
  }

  if (page?.image) {
    sources['1'] = page.image
  }

  return sources
}

function sortedImageKeys(imageSources: ImageSources) {
  return Object.keys(imageSources).sort((first, second) => Number(first) - Number(second))
}

function buildBodyContentItems(runs?: RichTextRun[], fallback = ''): BodyContentItem[] {
  const sourceRuns =
    runs && runs.length > 0
      ? runs
      : fallback
        ? [
            {
              text: fallback,
            },
          ]
        : []

  if (sourceRuns.length === 0) {
    return []
  }

  const fullText = sourceRuns.map((run) => run.text).join('')
  const items: BodyContentItem[] = []
  let cursor = 0
  let shouldTrimLeadingBreak = false

  INLINE_IMAGE_TOKEN.lastIndex = 0
  let match = INLINE_IMAGE_TOKEN.exec(fullText)

  while (match) {
    const matchStart = match.index
    const matchEnd = matchStart + match[0].length

    if (cursor < matchStart) {
      let textRuns = sliceRunsByRange(sourceRuns, cursor, matchStart)
      if (shouldTrimLeadingBreak) {
        textRuns = trimLeadingLineBreak(textRuns)
      }
      textRuns = trimTrailingLineBreak(textRuns)
      if (textRuns.length > 0) {
        items.push({
          type: 'text',
          runs: textRuns,
        })
      }
    }

    items.push({
      type:
        match[1].toLowerCase() === 'images' ||
        (parseInlineImageKeys(match[2])?.length ?? 0) > 1
          ? 'imageRow'
          : 'image',
      options:
        match[1].toLowerCase() === 'images' ||
        (parseInlineImageKeys(match[2])?.length ?? 0) > 1
          ? {
              imageKeys: parseInlineImageKeys(match[2]),
              ...parseInlineImageOptions(match[3]),
            }
          : {
              imageKey: parseInlineImageKeys(match[2])?.[0] ?? '1',
              ...parseInlineImageOptions(match[3]),
            },
    })

    cursor = matchEnd
    shouldTrimLeadingBreak = true
    match = INLINE_IMAGE_TOKEN.exec(fullText)
  }

  if (cursor < fullText.length) {
    let textRuns = sliceRunsByRange(sourceRuns, cursor, fullText.length)
    if (shouldTrimLeadingBreak) {
      textRuns = trimLeadingLineBreak(textRuns)
    }
    if (textRuns.length > 0) {
      items.push({
        type: 'text',
        runs: textRuns,
      })
    }
  }

  return items
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

function PageImage({
  source,
  width,
  height,
  align,
  fit,
  compact,
}: {
  source?: string
  width?: string
  height?: string
  align?: string
  fit?: string
  compact?: boolean
}) {
  const [failed, setFailed] = useState(false)
  const resolvedSource = resolveImageSource(source)

  if (!resolvedSource || failed) {
    return null
  }

  return (
    <div
      className={`pageImageFrame${compact ? ' pageImageFrame-compact' : ''}`}
      style={{ justifyContent: resolveImageAlign(align) }}
    >
      <img
        className="pageImage"
        src={resolvedSource}
        alt=""
        loading="lazy"
        style={{
          width: normalizeCssSize(width),
          height: normalizeCssSize(height),
          objectFit: resolveImageFit(fit),
        }}
        onError={() => setFailed(true)}
      />
    </div>
  )
}

function PageImageRow({
  imageKeys,
  imageSources,
  width,
  height,
  align,
  fit,
  compact,
}: {
  imageKeys?: string[]
  imageSources: ImageSources
  width?: string
  height?: string
  align?: string
  fit?: string
  compact?: boolean
}) {
  const resolvedKeys = (imageKeys && imageKeys.length > 0
    ? imageKeys
    : sortedImageKeys(imageSources)
  ).filter((imageKey) => Boolean(imageSources[imageKey]))

  if (resolvedKeys.length === 0) {
    return null
  }

  return (
    <div className={`pageImageRow${compact ? ' pageImageRow-compact' : ''}`}>
      {resolvedKeys.map((imageKey) => (
        <div key={imageKey} className="pageImageRowItem">
          <PageImage
            source={imageSources[imageKey]}
            width={width}
            height={height}
            align={align}
            fit={fit}
            compact={compact}
          />
        </div>
      ))}
    </div>
  )
}

function BodyTextBlock({
  runs,
  multiplier,
  defaultFontFamily,
  cellFontFamily,
  baseFontSize,
  baseTextColor,
}: {
  runs: RichTextRun[]
  multiplier: number
  defaultFontFamily?: string
  cellFontFamily?: string
  baseFontSize?: number
  baseTextColor?: string
}) {
  if (runs.length === 0) {
    return null
  }

  return (
    <RichText
      runs={runs}
      fallback=""
      className="pageBodyText"
      multiplier={multiplier}
      defaultFontFamily={defaultFontFamily}
      cellFontFamily={cellFontFamily}
      baseFontSize={baseFontSize}
      baseTextColor={baseTextColor}
    />
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
  const imageSources = buildImageSourceMap(page)
  const imageKeys = sortedImageKeys(imageSources)
  const bodyItems = buildBodyContentItems(page?.bodyRuns, page?.body || '')
  const hasInlineImage = bodyItems.some(
    (item) => item.type === 'image' || item.type === 'imageRow',
  )
  const hasBodyText = bodyItems.some((item) => item.type === 'text')
  const imagePlacement = resolveImagePosition(page?.imagePosition)
  const defaultImageItems: BodyContentItem[] = imageKeys.map((imageKey) => ({
    type: 'image',
    options: { imageKey },
  }))
  const bodyContentItems: BodyContentItem[] =
    hasInlineImage || imageKeys.length === 0
      ? bodyItems
      : imagePlacement === 'top'
        ? [
            ...defaultImageItems,
            ...bodyItems,
          ]
        : [
            ...bodyItems,
            ...defaultImageItems,
          ]

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
        <div className="pageBodyContent">
          {bodyContentItems.map((item, itemIndex) =>
            item.type === 'text' ? (
              <BodyTextBlock
                key={`text-${itemIndex}`}
                runs={item.runs}
                multiplier={bodyMultiplier}
                defaultFontFamily={settings?.bodyFontFamily}
                cellFontFamily={bodyStyle?.fontFamily}
                baseFontSize={bodyStyle?.fontSize}
                baseTextColor={bodyStyle?.textColor}
              />
            ) : (
              item.type === 'image' ? (
                <PageImage
                  key={`image-${itemIndex}`}
                  source={item.options.source ?? imageSources[item.options.imageKey ?? '1']}
                  width={item.options.width ?? page.imageWidth}
                  height={item.options.height ?? page.imageHeight}
                  align={item.options.align ?? page.imageAlign}
                  fit={item.options.fit ?? page.imageFit}
                  compact={hasBodyText}
                />
              ) : (
                <PageImageRow
                  key={`image-row-${itemIndex}`}
                  imageKeys={item.options.imageKeys}
                  imageSources={imageSources}
                  width={item.options.width ?? page.imageWidth}
                  height={item.options.height ?? page.imageHeight}
                  align={item.options.align ?? page.imageAlign}
                  fit={item.options.fit ?? page.imageFit}
                  compact={hasBodyText}
                />
              )
            ),
          )}
        </div>
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
          Put <code>{'{{image}}'}</code> inside <code>body</code> to place the image
          between text blocks. Use <code>{'{{image:2}}'}</code> for <code>image_2</code>.
          Use <code>{'{{images:1,2,3}}'}</code> to place multiple images in one row.
          Example: <code>{'{{images:1,2 width=160px align=center}}'}</code>
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
