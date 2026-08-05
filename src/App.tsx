import { useEffect, useRef, useState, type CSSProperties, type ReactNode } from 'react'
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
  type PdfMetadata,
  type RenderSettings,
  type RichTextRun,
  type SideBlockDefinition,
  type SideBlockDefinitions,
} from './lib/hintbook'

const defaultAppsScriptSource =
  'https://script.google.com/macros/s/AKfycbxujRiMhpqRckPyMcnWahn4eJ6cvl3SVG5rGGao7_55iWpjKq5duNBHIeVtH4MyzifmFw/exec'

const BASE_FONT_SIZE = 10
const STEP_FONT_SCALE = 4
const BODY_FONT_SCALE = 2.5
const INLINE_BODY_TOKEN =
  /\{\{(?:(images?|img)(?::([0-9,\s]+))?(?:\s+([^}]+))?|space\s*:\s*((?:\d+(?:\.\d+)?|\.\d+)(?:px|mm|cm|pt|em|rem)?)\s*)\}\}/gi
const INLINE_DECORATION_TOKEN = /\[\[(\/?)(hl|box)(?:\s*:\s*([^[\]]+))?\]\]/gi
const DEFAULT_INLINE_HIGHLIGHT_COLOR = '#fff2a8'
const DEFAULT_INLINE_BOX_COLOR = '#d94b67'
const DRIVE_IMAGE_MAX_WIDTH = 4000
const SHEET_SOURCE_STORAGE_KEY = 'hintbookmaker.sheetSource'
const APPS_SCRIPT_SOURCE_STORAGE_KEY = 'hintbookmaker.appsScriptSource'
const PREVIEW_SCALE_STORAGE_KEY = 'hintbookmaker.previewScale'
const PREVIEW_SCALE_MIN = 55
const PREVIEW_SCALE_MAX = 100
const PREVIEW_SCALE_STEP = 5
const SHEET_CANVAS_WIDTH = 1122
const SHEET_CANVAS_HEIGHT = 794

type InlineImageOptions = {
  source?: string
  width?: string
  height?: string
  align?: string
  fit?: string
  imageKey?: string
  imageKeys?: string[]
  imageOptionsByKey?: Record<
    string,
    {
      source?: string
      width?: string
      height?: string
      align?: string
      fit?: string
    }
  >
}

type InlineBoxDecoration = {
  id: number
  backgroundColor: string
  textColor?: string
}

type RenderTextRun = RichTextRun & {
  inlineBox?: InlineBoxDecoration
}

type BodyContentItem =
  | {
      type: 'text'
      runs: RenderTextRun[]
    }
  | {
      type: 'image'
      options: InlineImageOptions
    }
  | {
      type: 'imageRow'
      options: InlineImageOptions
    }
  | {
      type: 'space'
      size: string
    }

type SideInlineDirection =
  | 'horizontal'
  | 'vertical'
  | 'upright'
  | 'tcy'
  | 'cw'
  | 'ccw'

type SideTextSegment = {
  direction?: SideInlineDirection
  runs: RichTextRun[]
}

type PreviewMode = 'print' | 'book'

type PreviewSpread = {
  key: string
  leftPage: FormattedPage | null
  rightPage: FormattedPage | null
  label: string
  meta: string
}

type ScrollNavDirection = 'up' | 'down'

function sanitizePdfFileNamePart(value?: string) {
  return Array.from((value ?? '').trim(), (character) =>
    character.charCodeAt(0) < 32 || '\\/:*?"<>|'.includes(character) ? '_' : character,
  )
    .join('')
    .replace(/[. ]+$/g, '')
}

function buildPdfTitle(previewMode: PreviewMode, metadata?: PdfMetadata) {
  const category = previewMode === 'book' ? '【冊子版】' : '【印刷用】'
  const name = sanitizePdfFileNamePart(metadata?.name)
  const version = sanitizePdfFileNamePart(metadata?.version)
  return `${category}${name}ヒント${version}`
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

function readStoredValue(storageKey: string, fallback: string) {
  if (typeof window === 'undefined') {
    return fallback
  }

  const storedValue = window.localStorage.getItem(storageKey)?.trim()
  return storedValue || fallback
}

function readStoredNumber(storageKey: string, fallback: number) {
  if (typeof window === 'undefined') {
    return fallback
  }

  const storedValue = Number(window.localStorage.getItem(storageKey))
  return Number.isFinite(storedValue) ? storedValue : fallback
}

function clampNumber(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), max)
}

function countConfiguredSettings(settings?: RenderSettings) {
  if (!settings) {
    return 0
  }

  return Object.values(settings).filter((value) => value !== undefined && value !== '')
    .length
}

function countSideDefinitions(sideDefinitions?: SideBlockDefinitions) {
  if (!sideDefinitions) {
    return 0
  }

  return Object.keys(sideDefinitions).length
}

function pageScrollTargetOffset() {
  const header = document.querySelector('.appHeader')
  return (header instanceof HTMLElement ? header.offsetHeight : 0) + 12
}

function scrollToPagePosition(top: number) {
  window.scrollTo({
    top: Math.max(top - pageScrollTargetOffset(), 0),
    behavior: 'smooth',
  })
}

function scrollToBoundary(position: 'top' | 'bottom') {
  window.scrollTo({
    top:
      position === 'top'
        ? 0
        : Math.max(document.documentElement.scrollHeight - window.innerHeight, 0),
    behavior: 'smooth',
  })
}

function scrollByPreviewPage(direction: ScrollNavDirection) {
  const sections = Array.from(document.querySelectorAll('.sheetPreview')).filter(
    (section): section is HTMLElement => section instanceof HTMLElement,
  )

  if (sections.length === 0) {
    window.scrollBy({
      top: direction === 'down' ? window.innerHeight * 0.85 : -window.innerHeight * 0.85,
      behavior: 'smooth',
    })
    return
  }

  const currentTop = window.scrollY + pageScrollTargetOffset()
  const sectionTops = sections.map((section) => ({
    section,
    top: section.getBoundingClientRect().top + window.scrollY,
  }))

  if (direction === 'down') {
    const nextSection = sectionTops.find(({ top }) => top > currentTop + 12)
    if (nextSection) {
      scrollToPagePosition(nextSection.top)
      return
    }

    scrollToBoundary('bottom')
    return
  }

  const previousSection = [...sectionTops].reverse().find(({ top }) => top < currentTop - 12)
  if (previousSection) {
    scrollToPagePosition(previousSection.top)
    return
  }

  scrollToBoundary('top')
}

function ScrollNavIcon({ direction }: { direction: 'top' | 'up' | 'down' | 'bottom' }) {
  const pathByDirection = {
    top: 'M12 5l-5 5h3v6h4v-6h3l-5-5zm0 8l-5 5h3v1h4v-1h3l-5-5z',
    up: 'M12 7l-5 5h3v5h4v-5h3l-5-5z',
    down: 'M12 17l5-5h-3V7h-4v5H7l5 5z',
    bottom: 'M12 19l5-5h-3V8h-4v6H7l5 5zm0-8l5-5h-3V5h-4v1H7l5 5z',
  } as const

  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
      <path d={pathByDirection[direction]} />
    </svg>
  )
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
    return `https://drive.google.com/thumbnail?id=${driveFileId}&sz=w${DRIVE_IMAGE_MAX_WIDTH}`
  }

  try {
    return new URL(trimmed).toString()
  } catch {
    return null
  }
}

function waitForNextFrame() {
  return new Promise<void>((resolve) => {
    window.requestAnimationFrame(() => resolve())
  })
}

function waitForImageElement(image: HTMLImageElement) {
  const decodeImage = () =>
    typeof image.decode === 'function'
      ? image.decode().catch(() => undefined)
      : Promise.resolve()

  if (image.complete) {
    return decodeImage()
  }

  return new Promise<void>((resolve) => {
    const finish = () => {
      cleanup()
      void decodeImage().finally(resolve)
    }

    const cleanup = () => {
      image.removeEventListener('load', finish)
      image.removeEventListener('error', finish)
    }

    image.addEventListener('load', finish, { once: true })
    image.addEventListener('error', finish, { once: true })
  })
}

async function waitForPrintableImages(root: ParentNode = document) {
  const images = Array.from(root.querySelectorAll('img'))
  if (images.length === 0) {
    return
  }

  await Promise.all(images.map((image) => waitForImageElement(image)))
  await waitForNextFrame()
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

function cloneRunWithText<T extends RichTextRun>(run: T, text: string): T {
  return {
    ...run,
    text,
  } as T
}

function runsHaveSameStyle(first: RichTextRun, second: RichTextRun) {
  const firstBox = (first as RenderTextRun).inlineBox
  const secondBox = (second as RenderTextRun).inlineBox

  return (
    first.backgroundColor === second.backgroundColor &&
    first.textColor === second.textColor &&
    first.fontFamily === second.fontFamily &&
    first.fontSize === second.fontSize &&
    first.bold === second.bold &&
    first.italic === second.italic &&
    first.underline === second.underline &&
    first.strikethrough === second.strikethrough &&
    firstBox?.id === secondBox?.id
  )
}

function appendRuns<T extends RichTextRun>(target: T[], runs: T[]) {
  for (const run of runs) {
    const lastRun = target[target.length - 1]

    if (lastRun && runsHaveSameStyle(lastRun, run)) {
      lastRun.text += run.text
      continue
    }

    target.push({ ...run })
  }
}

function sliceRunsByRange<T extends RichTextRun>(
  runs: T[],
  startIndex: number,
  endIndex: number,
) {
  if (startIndex >= endIndex) {
    return []
  }

  const slicedRuns: T[] = []
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

function trimLeadingLineBreak<T extends RichTextRun>(runs: T[]) {
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

function trimTrailingLineBreak<T extends RichTextRun>(runs: T[]) {
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

function resolveInlineHighlightColor(value?: string) {
  const trimmed = value?.trim()
  return trimmed || DEFAULT_INLINE_HIGHLIGHT_COLOR
}

function parseHexColor(value: string) {
  const match = value.trim().match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i)
  if (!match) {
    return undefined
  }

  const hex =
    match[1].length === 3
      ? match[1]
          .split('')
          .map((digit) => digit + digit)
          .join('')
      : match[1]

  return {
    red: Number.parseInt(hex.slice(0, 2), 16),
    green: Number.parseInt(hex.slice(2, 4), 16),
    blue: Number.parseInt(hex.slice(4, 6), 16),
  }
}

function resolveInlineBoxTextColor(backgroundColor: string) {
  const color = parseHexColor(backgroundColor)
  if (!color) {
    return undefined
  }

  const perceivedBrightness =
    (color.red * 299 + color.green * 587 + color.blue * 114) / 1000
  return perceivedBrightness >= 160 ? '#1a1a1a' : '#ffffff'
}

function resolveInlineBox(value: string | undefined, id: number): InlineBoxDecoration {
  const backgroundColor = value?.trim() || DEFAULT_INLINE_BOX_COLOR
  return {
    id,
    backgroundColor,
    textColor: resolveInlineBoxTextColor(backgroundColor),
  }
}

function applyInlineDecorations(sourceRuns: RichTextRun[]): RenderTextRun[] {
  if (sourceRuns.length === 0) {
    return sourceRuns
  }

  const fullText = sourceRuns.map((run) => run.text).join('')
  if (!/\[\[\/?(?:hl|box)(?:\s*:|\]\])/i.test(fullText)) {
    return sourceRuns
  }

  const decoratedRuns: RenderTextRun[] = []
  const activeHighlightColors: string[] = []
  const activeBoxes: InlineBoxDecoration[] = []
  let nextBoxId = 1
  let cursor = 0

  INLINE_DECORATION_TOKEN.lastIndex = 0
  let match = INLINE_DECORATION_TOKEN.exec(fullText)

  while (match) {
    const matchStart = match.index
    const matchEnd = matchStart + match[0].length

    if (cursor < matchStart) {
      const textRuns: RenderTextRun[] = sliceRunsByRange(
        sourceRuns,
        cursor,
        matchStart,
      ).map((run) => ({
        ...run,
        backgroundColor:
          activeHighlightColors[activeHighlightColors.length - 1] ?? run.backgroundColor,
        inlineBox: activeBoxes[activeBoxes.length - 1],
      }))
      appendRuns(decoratedRuns, textRuns)
    }

    const decorationType = match[2].toLowerCase()
    if (match[1] === '/') {
      if (decorationType === 'hl') {
        activeHighlightColors.pop()
      } else {
        activeBoxes.pop()
      }
    } else if (decorationType === 'hl') {
      activeHighlightColors.push(resolveInlineHighlightColor(match[3]))
    } else {
      activeBoxes.push(resolveInlineBox(match[3], nextBoxId))
      nextBoxId += 1
    }

    cursor = matchEnd
    match = INLINE_DECORATION_TOKEN.exec(fullText)
  }

  if (cursor < fullText.length) {
    const textRuns: RenderTextRun[] = sliceRunsByRange(
      sourceRuns,
      cursor,
      fullText.length,
    ).map((run) => ({
      ...run,
      backgroundColor:
        activeHighlightColors[activeHighlightColors.length - 1] ?? run.backgroundColor,
      inlineBox: activeBoxes[activeBoxes.length - 1],
    }))
    appendRuns(decoratedRuns, textRuns)
  }

  return decoratedRuns
}

function applyInlineImageOption(
  options: Pick<InlineImageOptions, 'source' | 'width' | 'height' | 'align' | 'fit'>,
  key: string,
  value: string,
) {
  if (key === 'src' || key === 'source' || key === 'url') {
    options.source = value
  } else if (key === 'width' || key === 'w') {
    options.width = value
  } else if (key === 'height' || key === 'h') {
    options.height = value
  } else if (key === 'align') {
    options.align = value
  } else if (key === 'fit') {
    options.fit = value
  }
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

    const perImageMatch = key.match(/^(src|source|url|width|w|height|h|align|fit)_?(\d+)$/)
    if (perImageMatch) {
      const imageKey = perImageMatch[2]
      const imageOptions = options.imageOptionsByKey?.[imageKey] ?? {}
      applyInlineImageOption(imageOptions, perImageMatch[1], rawValue)
      options.imageOptionsByKey = {
        ...(options.imageOptionsByKey ?? {}),
        [imageKey]: imageOptions,
      }
      continue
    }

    if (key === 'index' || key === 'image' || key === 'key') {
      options.imageKey = rawValue
    } else {
      applyInlineImageOption(options, key, rawValue)
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

  const normalizedRuns = applyInlineDecorations(sourceRuns)
  const fullText = normalizedRuns.map((run) => run.text).join('')
  const items: BodyContentItem[] = []
  let cursor = 0
  let shouldTrimLeadingBreak = false

  INLINE_BODY_TOKEN.lastIndex = 0
  let match = INLINE_BODY_TOKEN.exec(fullText)

  while (match) {
    const matchStart = match.index
    const matchEnd = matchStart + match[0].length

    if (cursor < matchStart) {
      let textRuns = sliceRunsByRange(normalizedRuns, cursor, matchStart)
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

    const spaceSize = normalizeCssSize(match[4]?.toLowerCase())
    if (spaceSize) {
      items.push({
        type: 'space',
        size: spaceSize,
      })
    } else {
      const imageType = match[1]?.toLowerCase() ?? 'image'
      const imageKeys = parseInlineImageKeys(match[2])
      const isImageRow = imageType === 'images' || (imageKeys?.length ?? 0) > 1

      items.push({
        type: isImageRow ? 'imageRow' : 'image',
        options: isImageRow
          ? {
              imageKeys,
              ...parseInlineImageOptions(match[3]),
            }
          : {
              imageKey: imageKeys?.[0] ?? '1',
              ...parseInlineImageOptions(match[3]),
            },
      })
    }

    cursor = matchEnd
    shouldTrimLeadingBreak = true
    match = INLINE_BODY_TOKEN.exec(fullText)
  }

  if (cursor < fullText.length) {
    let textRuns = sliceRunsByRange(normalizedRuns, cursor, fullText.length)
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
    backgroundColor: run.backgroundColor,
    color: run.textColor ?? options.baseTextColor,
    fontFamily: options.omitFontFamily
      ? undefined
      : resolveRunFontFamily(run, options.defaultFontFamily, options.cellFontFamily),
    fontSize: scaledFontSize(run.fontSize ?? options.baseFontSize, multiplier),
    fontWeight: run.bold === undefined ? undefined : run.bold ? '700' : '400',
    fontStyle: run.italic === undefined ? undefined : run.italic ? 'italic' : 'normal',
    textDecoration: resolveTextDecoration(run.underline, run.strikethrough),
    padding: run.backgroundColor ? '0.02em 0.18em' : undefined,
  }
}

type RichTextRunGroup = {
  inlineBox?: InlineBoxDecoration
  runs: RenderTextRun[]
}

function groupRunsByInlineBox(runs: RichTextRun[]) {
  const groups: RichTextRunGroup[] = []

  for (const sourceRun of runs) {
    const run = sourceRun as RenderTextRun
    const lastGroup = groups[groups.length - 1]

    if (lastGroup && lastGroup.inlineBox?.id === run.inlineBox?.id) {
      lastGroup.runs.push(run)
      continue
    }

    groups.push({
      inlineBox: run.inlineBox,
      runs: [run],
    })
  }

  return groups
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

  const runGroups = groupRunsByInlineBox(sourceRuns)

  return (
    <div className={className}>
      {runGroups.map((group, groupIndex) => {
        const effectiveBaseTextColor = group.inlineBox?.textColor ?? baseTextColor
        const content = group.runs.map((run, runIndex) => {
          const parts = run.text.split('\n')
          const style = runStyle(run, multiplier, {
            defaultFontFamily,
            cellFontFamily,
            baseFontSize,
            baseTextColor: effectiveBaseTextColor,
            omitFontFamily,
          })

          return (
            <span key={`${groupIndex}-${runIndex}-${run.text}`}>
              {parts.map((part, partIndex) => (
                <span key={`${runIndex}-${partIndex}`} style={style}>
                  {part}
                  {partIndex < parts.length - 1 ? <br /> : null}
                </span>
              ))}
            </span>
          )
        })

        if (!group.inlineBox) {
          return <span key={`plain-${groupIndex}`}>{content}</span>
        }

        const firstRun = group.runs[0]
        return (
          <span
            key={`box-${group.inlineBox.id}`}
            className="inlineTextBox"
            style={{
              backgroundColor: group.inlineBox.backgroundColor,
              color: effectiveBaseTextColor,
              fontSize: scaledFontSize(firstRun?.fontSize ?? baseFontSize, multiplier),
            }}
          >
            {content}
          </span>
        )
      })}
    </div>
  )
}

function buildSideFallbackRuns(definition: SideBlockDefinition) {
  if (definition.textRuns && definition.textRuns.length > 0) {
    return definition.textRuns
  }

  if (!definition.text) {
    return undefined
  }

  return [
    {
      text: definition.text,
      bold: definition.bold,
      italic: definition.italic,
      underline: definition.underline,
      strikethrough: definition.strikethrough,
    },
  ]
}

function appendSideTextSegment(
  segments: SideTextSegment[],
  direction: SideInlineDirection | undefined,
  runs: RichTextRun[],
) {
  if (runs.length === 0) {
    return
  }

  const lastSegment = segments[segments.length - 1]
  if (lastSegment && lastSegment.direction === direction) {
    appendRuns(lastSegment.runs, runs)
    return
  }

  segments.push({
    direction,
    runs: runs.map((run) => ({ ...run })),
  })
}

function normalizeSideInlineDirectionToken(token: string): SideInlineDirection {
  const normalized = token.trim().toLowerCase()
  if (normalized === 'h') return 'horizontal'
  if (normalized === 'v') return 'vertical'
  if (normalized === 'u') return 'upright'
  return normalized as SideInlineDirection
}

function parseSideTextSegments(runs?: RichTextRun[]) {
  if (!runs || runs.length === 0) {
    return null
  }

  const fullText = runs.map((run) => run.text).join('')
  const tokenPattern = /\[(\/?)\s*(h|v|u|upright|tcy|cw|ccw)\s*\]/gi
  const directionStack: SideInlineDirection[] = []
  const segments: SideTextSegment[] = []
  let cursor = 0
  let sawToken = false
  let match = tokenPattern.exec(fullText)

  while (match) {
    sawToken = true
    const matchStart = match.index
    const matchEnd = matchStart + match[0].length

    appendSideTextSegment(
      segments,
      directionStack[directionStack.length - 1],
      sliceRunsByRange(runs, cursor, matchStart),
    )

    const direction = normalizeSideInlineDirectionToken(match[2])
    if (match[1] === '/') {
      if (directionStack[directionStack.length - 1] !== direction) {
        return null
      }
      directionStack.pop()
    } else {
      directionStack.push(direction)
    }

    cursor = matchEnd
    match = tokenPattern.exec(fullText)
  }

  if (!sawToken || directionStack.length > 0) {
    return null
  }

  appendSideTextSegment(
    segments,
    undefined,
    sliceRunsByRange(runs, cursor, fullText.length),
  )

  return segments
}

function resolveSideBlockBaseDirection(
  definition: SideBlockDefinition,
): Exclude<SideInlineDirection, 'tcy'> {
  if (definition.textDirection === 'upright') return 'upright'
  if (definition.textDirection === 'vertical') return 'vertical'
  if (definition.textDirection === 'rotateCw') return 'cw'
  if (definition.textDirection === 'rotateCcw') return 'ccw'
  if (definition.textDirection === 'horizontal') return 'horizontal'
  if (definition.textRotation === 'cw') return 'cw'
  if (definition.textRotation === 'ccw') return 'ccw'
  return 'horizontal'
}

function resolveSideBlockTextRotationClass(definition: SideBlockDefinition) {
  if (definition.textDirection === 'upright') {
    return 'pageSideBlockText-upright'
  }

  if (definition.textDirection === 'vertical') {
    return 'pageSideBlockText-vertical'
  }

  if (definition.textDirection === 'rotateCw') {
    return 'pageSideBlockText-rotateCw'
  }

  if (definition.textDirection === 'rotateCcw') {
    return 'pageSideBlockText-rotateCcw'
  }

  if (definition.textDirection === 'horizontal') {
    return ''
  }

  if (definition.textRotation === 'cw') {
    return 'pageSideBlockText-rotateCw'
  }

  if (definition.textRotation === 'ccw') {
    return 'pageSideBlockText-rotateCcw'
  }

  return ''
}

function resolveSideTextSegmentClass(direction: SideInlineDirection) {
  if (direction === 'upright') return 'pageSideTextSegment-upright'
  if (direction === 'vertical') return 'pageSideTextSegment-vertical'
  if (direction === 'tcy') return 'pageSideTextSegment-tcy'
  if (direction === 'cw') return 'pageSideTextSegment-cw'
  if (direction === 'ccw') return 'pageSideTextSegment-ccw'
  return 'pageSideTextSegment-horizontal'
}

function SideBlockText({
  definition,
  position,
  defaultFontFamily,
}: {
  definition: SideBlockDefinition
  position: 'left' | 'right'
  defaultFontFamily?: string
}) {
  const runs = buildSideFallbackRuns(definition)
  const segments = parseSideTextSegments(runs)

  if (!segments) {
    const directionClass = resolveSideBlockTextRotationClass(definition)
    return (
      <RichText
        runs={runs}
        fallback=""
        className={`pageSideBlockText pageSideBlockText-${position}${
          directionClass ? ` ${directionClass}` : ''
        }`}
        multiplier={1}
        defaultFontFamily={defaultFontFamily}
        cellFontFamily={defaultFontFamily}
        baseFontSize={definition.fontSize}
        baseTextColor={
          definition.textColor ?? (definition.backgroundColor ? '#ffffff' : undefined)
        }
      />
    )
  }

  const baseDirection = resolveSideBlockBaseDirection(definition)

  return (
    <div className="pageSideTextSegments">
      {segments.map((segment, segmentIndex) => {
        const direction = segment.direction ?? baseDirection
        return (
          <RichText
            key={`${segmentIndex}-${direction}`}
            runs={segment.runs}
            fallback=""
            className={`pageSideTextSegment ${resolveSideTextSegmentClass(direction)}`}
            multiplier={1}
            defaultFontFamily={defaultFontFamily}
            cellFontFamily={defaultFontFamily}
            baseFontSize={definition.fontSize}
            baseTextColor={
              definition.textColor ?? (definition.backgroundColor ? '#ffffff' : undefined)
            }
          />
        )
      })}
    </div>
  )
}

function hasExplicitImageSize(width?: string, height?: string) {
  return Boolean(normalizeCssSize(width) || normalizeCssSize(height))
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
  const [failedSource, setFailedSource] = useState<string | null>(null)
  const [portraitSource, setPortraitSource] = useState<string | null>(null)
  const resolvedSource = resolveImageSource(source)
  const isPortrait = resolvedSource !== null && portraitSource === resolvedSource
  const hasExplicitSize = hasExplicitImageSize(width, height)

  if (!resolvedSource || failedSource === resolvedSource) {
    return null
  }

  return (
    <div
      className={`pageImageFrame${
        compact && !hasExplicitSize ? ' pageImageFrame-compact' : ''
      }${compact && !hasExplicitSize && isPortrait ? ' pageImageFrame-compactPortrait' : ''}`}
      style={{ justifyContent: resolveImageAlign(align) }}
    >
      <img
        className="pageImage"
        src={resolvedSource}
        alt=""
        loading="eager"
        fetchPriority="high"
        style={{
          width: normalizeCssSize(width),
          height: normalizeCssSize(height),
          objectFit: resolveImageFit(fit),
        }}
        onLoad={(event) => {
          const image = event.currentTarget
          setPortraitSource(
            image.naturalHeight > image.naturalWidth * 1.05 ? resolvedSource : null,
          )
        }}
        onError={() => setFailedSource(resolvedSource)}
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
  imageOptionsByKey,
}: {
  imageKeys?: string[]
  imageSources: ImageSources
  width?: string
  height?: string
  align?: string
  fit?: string
  compact?: boolean
  imageOptionsByKey?: InlineImageOptions['imageOptionsByKey']
}) {
  const resolvedKeys = (imageKeys && imageKeys.length > 0
    ? imageKeys
    : sortedImageKeys(imageSources)
  ).filter((imageKey) => Boolean(imageSources[imageKey]))
  const hasExplicitRowSize =
    hasExplicitImageSize(width, height) ||
    resolvedKeys.some((imageKey) =>
      hasExplicitImageSize(
        imageOptionsByKey?.[imageKey]?.width,
        imageOptionsByKey?.[imageKey]?.height,
      ),
    )

  if (resolvedKeys.length === 0) {
    return null
  }

  return (
    <div
      className={`pageImageRow${compact && !hasExplicitRowSize ? ' pageImageRow-compact' : ''}`}
    >
      {resolvedKeys.map((imageKey) => {
        const itemOptions = imageOptionsByKey?.[imageKey]
        const resolvedWidth = normalizeCssSize(itemOptions?.width ?? width)

        return (
        <div
          key={imageKey}
          className="pageImageRowItem"
          style={
            resolvedWidth
              ? {
                  flex: `0 0 ${resolvedWidth}`,
                  width: resolvedWidth,
                }
              : undefined
          }
        >
          <PageImage
            source={itemOptions?.source ?? imageSources[imageKey]}
            width={itemOptions?.width ?? width}
            height={itemOptions?.height ?? height}
            align={itemOptions?.align ?? align}
            fit={itemOptions?.fit ?? fit}
            compact={compact}
          />
        </div>
        )
      })}
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

type ResolvedSideBlock = {
  definition: SideBlockDefinition
  height: number
}

function parseSideReferences(value?: string) {
  if (!value) {
    return []
  }

  return value
    .split(/[,\n\r\u3001\uFF0C\s]+/)
    .map((entry) => entry.trim())
    .filter(Boolean)
}

function resolveSideBlocks(
  sideValue: string | undefined,
  sideDefinitions?: SideBlockDefinitions,
): ResolvedSideBlock[] {
  if (!sideDefinitions) {
    return []
  }

  const definitions = parseSideReferences(sideValue)
    .map((id) => sideDefinitions[id])
    .filter((definition): definition is SideBlockDefinition => Boolean(definition))

  if (definitions.length === 0) {
    return []
  }

  const specifiedHeight = definitions.reduce(
    (total, definition) => total + (definition.height ?? 0),
    0,
  )
  const missingHeightCount = definitions.filter(
    (definition) => definition.height === undefined,
  ).length

  if (specifiedHeight > 100) {
    return definitions
      .map((definition) => ({
        definition,
        height: ((definition.height ?? 0) / specifiedHeight) * 100,
      }))
      .filter((entry) => entry.height > 0)
  }

  const remainingHeight = Math.max(100 - specifiedHeight, 0)
  const defaultHeight = missingHeightCount > 0 ? remainingHeight / missingHeightCount : 0

  return definitions.map((definition) => ({
    definition,
    height: definition.height ?? defaultHeight,
  }))
}

function buildBookPreviewSpreads(pages: FormattedPage[]): PreviewSpread[] {
  if (pages.length === 0) {
    return []
  }

  const spreads: PreviewSpread[] = [
    {
      key: 'book-cover-open',
      leftPage: null,
      rightPage: pages[0] ?? null,
      label: 'book spread 1',
      meta: 'blank, 1',
    },
  ]

  for (let index = 1; index < pages.length; index += 2) {
    const leftPage = pages[index] ?? null
    const rightPage = pages[index + 1] ?? null
    const spreadNumber = spreads.length + 1
    const labels = [leftPage?.pageNo, rightPage?.pageNo].filter(Boolean)

    spreads.push({
      key: `book-${spreadNumber}-${leftPage?.order ?? 'blank'}-${rightPage?.order ?? 'blank'}`,
      leftPage,
      rightPage,
      label: `book spread ${spreadNumber}`,
      meta: labels.length > 0 ? labels.join(', ') : 'blank',
    })
  }

  return spreads
}

function PageSideLabel({
  page,
  position,
  defaultFontFamily,
  sideDefinitions,
}: {
  page: FormattedPage | null
  position: 'left' | 'right'
  defaultFontFamily?: string
  sideDefinitions?: SideBlockDefinitions
}) {
  const resolvedSideBlocks = resolveSideBlocks(page?.side, sideDefinitions)

  if (!page) {
    return <div className={`pageSideLabel pageSideLabel-${position}`} />
  }

  if (resolvedSideBlocks.length === 0) {
    return (
      <div className={`pageSideLabel pageSideLabel-${position}`}>
        <div
          className={`pageSideFallback pageSideFallback-${position}`}
          style={{ fontFamily: defaultFontFamily }}
        >
          {page.side || ''}
        </div>
      </div>
    )
  }

  const totalHeight = resolvedSideBlocks.reduce((sum, block) => sum + block.height, 0)
  const remainingHeight = Math.max(100 - totalHeight, 0)

  return (
    <div className={`pageSideLabel pageSideLabel-${position}`}>
      <div className="pageSideBlocks">
        {resolvedSideBlocks.map(({ definition, height }, blockIndex) => (
          (() => {
            const effectiveSideFontFamily = defaultFontFamily ?? definition.fontFamily

            return (
          <div
            key={`${definition.id}-${blockIndex}`}
            className="pageSideBlock"
            style={{
              flex: `0 0 ${height}%`,
              backgroundColor: definition.backgroundColor,
            }}
          >
            <div className="pageSideBlockContent">
              <div className="pageSideBlockTextFrame">
                <SideBlockText
                  definition={definition}
                  position={position}
                  defaultFontFamily={effectiveSideFontFamily}
                />
              </div>
            </div>
          </div>
            )
          })()
        ))}
        {remainingHeight > 0 ? (
          <div className="pageSideBlockSpacer" style={{ flex: `0 0 ${remainingHeight}%` }} />
        ) : null}
      </div>
    </div>
  )
}

function PagePreview({
  page,
  position,
  layoutPosition = position,
  settings,
  sideDefinitions,
}: {
  page: FormattedPage | null
  position: 'left' | 'right'
  layoutPosition?: 'left' | 'right'
  settings?: RenderSettings
  sideDefinitions?: SideBlockDefinitions
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
  const pageNumberStyle = {
    fontFamily: pageNoFontFamily,
  }
  const pageLayoutStyle = (
    settings?.sideWidth
      ? {
          '--page-side-width': normalizeCssSize(settings.sideWidth),
        }
      : undefined
  ) as CSSProperties | undefined
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
      <article
        className={`pagePreview pagePreview-${position}`}
        style={pageLayoutStyle}
      >
        <div className="pageStep" />
        <div className={`pageSideLabel pageSideLabel-${layoutPosition}`} />
        <div className={`pageBody pageBody-${layoutPosition}`} />
        <div className="pageNumber" />
      </article>
    )
  }

  return (
    <article className={`pagePreview pagePreview-${position}`} style={pageLayoutStyle}>
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
      <PageSideLabel
        page={page}
        position={layoutPosition}
        defaultFontFamily={sideFontFamily}
        sideDefinitions={sideDefinitions}
      />
      <div className={`pageBody pageBody-${layoutPosition}`} style={bodyBlockStyle}>
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
            ) : item.type === 'space' ? (
              <div
                key={`space-${itemIndex}`}
                className="pageBodySpacer"
                style={{ height: item.size }}
                aria-hidden="true"
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
                  imageOptionsByKey={item.options.imageOptionsByKey}
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

function ResponsiveSheetCanvas({
  children,
  displayScale,
}: {
  children: ReactNode
  displayScale: number
}) {
  const containerRef = useRef<HTMLDivElement | null>(null)
  const [scale, setScale] = useState(1)

  useEffect(() => {
    const container = containerRef.current
    if (!container) {
      return
    }

    const updateScale = () => {
      const nextScale =
        Math.min(container.clientWidth / SHEET_CANVAS_WIDTH, 1) * displayScale
      setScale(nextScale > 0 ? nextScale : 1)
    }

    updateScale()

    const observer = new ResizeObserver(() => {
      updateScale()
    })

    observer.observe(container)
    window.addEventListener('resize', updateScale)

    return () => {
      observer.disconnect()
      window.removeEventListener('resize', updateScale)
    }
  }, [displayScale])

  return (
    <div ref={containerRef} className="sheetViewport">
      <div
        className="sheetScaleBox"
        style={{
          width: `${SHEET_CANVAS_WIDTH * scale}px`,
          height: `${SHEET_CANVAS_HEIGHT * scale}px`,
        }}
      >
        <div
          className="sheetCanvas"
          style={{
            transform: `scale(${scale})`,
          }}
        >
          {children}
        </div>
      </div>
    </div>
  )
}

function App() {
  const [sheetSource, setSheetSource] = useState(() =>
    readStoredValue(SHEET_SOURCE_STORAGE_KEY, defaultSheetSource),
  )
  const [appsScriptSource, setAppsScriptSource] = useState(() =>
    readStoredValue(APPS_SCRIPT_SOURCE_STORAGE_KEY, defaultAppsScriptSource),
  )
  const [csvText, setCsvText] = useState(sampleCsv)
  const [result, setResult] = useState<FormatResult | null>(() =>
    formatHintBookFromCsv(sampleCsv),
  )
  const [status, setStatus] = useState('Ready.')
  const [loading, setLoading] = useState(false)
  const [previewMode, setPreviewMode] = useState<PreviewMode>('print')
  const [previewScalePercent, setPreviewScalePercent] = useState(() =>
    clampNumber(
      readStoredNumber(PREVIEW_SCALE_STORAGE_KEY, 80),
      PREVIEW_SCALE_MIN,
      PREVIEW_SCALE_MAX,
    ),
  )

  useEffect(() => {
    window.localStorage.setItem(SHEET_SOURCE_STORAGE_KEY, sheetSource)
  }, [sheetSource])

  useEffect(() => {
    window.localStorage.setItem(APPS_SCRIPT_SOURCE_STORAGE_KEY, appsScriptSource)
  }, [appsScriptSource])

  useEffect(() => {
    window.localStorage.setItem(
      PREVIEW_SCALE_STORAGE_KEY,
      String(previewScalePercent),
    )
  }, [previewScalePercent])

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
      const loadedSettingCount = countConfiguredSettings(next.settings)
      const loadedSideDefinitionCount = countSideDefinitions(next.sideDefinitions)
      setResult(next)
      setStatus(
        `Loaded Apps Script data and formatted ${next.pages.length} pages into ${next.spreads.length} print spreads. Shared settings: ${loadedSettingCount}. Side definitions: ${loadedSideDefinitionCount}.`,
      )
    } catch (error) {
      setStatus(error instanceof Error ? error.message : 'Apps Script load failed.')
    } finally {
      setLoading(false)
    }
  }

  const handleUseAppsScriptSample = () => {
    const next = formatHintBookFromAppsScript(sampleAppsScriptResponse)
    const loadedSettingCount = countConfiguredSettings(next.settings)
    const loadedSideDefinitionCount = countSideDefinitions(next.sideDefinitions)
    setResult(next)
    setStatus(
      `Loaded sample Apps Script data and formatted ${next.pages.length} pages into ${next.spreads.length} print spreads. Shared settings: ${loadedSettingCount}. Side definitions: ${loadedSideDefinitionCount}.`,
    )
  }

  const bookPreviewSpreads = result ? buildBookPreviewSpreads(result.pages) : []
  const previewSpreads: PreviewSpread[] =
    previewMode === 'book'
      ? bookPreviewSpreads
      : (result?.spreads.map((spread) => ({
          key: `print-${spread.sheetNumber}`,
          leftPage: spread.leftPage,
          rightPage: spread.rightPage,
          label: `sheet ${spread.sheetNumber}`,
          meta: `spread group ${spread.spreadNumber}`,
        })) ?? [])

  const handleExportPdf = async () => {
    if (!previewSpreads.length) {
      setStatus('No preview spreads available for PDF export.')
      return
    }

    setStatus('Preparing images for PDF export...')
    await waitForPrintableImages(document.querySelector('.previewArea') ?? document)

    const pdfTitle = buildPdfTitle(previewMode, result?.pdfMetadata)

    setStatus(
      `Opening print dialog for ${previewSpreads.length} ${
        previewMode === 'book' ? 'book preview' : 'print layout'
      } sheets. Choose "Save as PDF" to download as ${pdfTitle}.pdf.`,
    )
    const originalTitle = document.title
    document.title = pdfTitle
    try {
      window.print()
    } finally {
      document.title = originalTitle
    }
  }

  return (
    <main className="app">
      <header className="appHeader">
        <div className="appHeaderBar">
          <label style={{ display: 'none' }} aria-hidden="true">
            Google Sheets URL or ID
            <input
              type="text"
              value={sheetSource}
              onChange={(event) => setSheetSource(event.target.value)}
              tabIndex={-1}
            />
          </label>

          <button
            type="button"
            onClick={handleLoadSpreadsheet}
            disabled={loading}
            style={{ display: 'none' }}
            aria-hidden="true"
            tabIndex={-1}
          >
            {loading ? 'Loading...' : 'Load spreadsheet'}
          </button>
          <button
            type="button"
            onClick={() => setCsvText(sampleCsv)}
            style={{ display: 'none' }}
            aria-hidden="true"
            tabIndex={-1}
          >
            Use sample CSV
          </button>
          <button
            type="button"
            onClick={() => result && downloadJson(result)}
            disabled={!result}
            style={{ display: 'none' }}
            aria-hidden="true"
            tabIndex={-1}
          >
            Download JSON
          </button>
          <button
            type="button"
            onClick={handleUseAppsScriptSample}
            style={{ display: 'none' }}
            aria-hidden="true"
            tabIndex={-1}
          >
            Use Apps Script sample
          </button>

          <h1 className="appTitle">Hint Book Formatter</h1>
          <label className="appHeaderField">
            <span className="srOnly">Apps Script Web App URL</span>
            <input
              type="text"
              value={appsScriptSource}
              onChange={(event) => setAppsScriptSource(event.target.value)}
              placeholder="Apps Script Web App URL"
            />
          </label>
          <button type="button" onClick={handleLoadAppsScript} disabled={loading}>
            {loading ? 'Loading...' : 'Load Apps Script'}
          </button>
          <button
            type="button"
            onClick={handleExportPdf}
            disabled={!previewSpreads.length}
          >
            Export PDF
          </button>
          <div className="buttonRow appHeaderActions">
            <button
              type="button"
              onClick={() => setPreviewMode('print')}
              disabled={previewMode === 'print'}
            >
              印刷用
            </button>
            <button
              type="button"
              onClick={() => setPreviewMode('book')}
              disabled={previewMode === 'book'}
            >
              冊子版
            </button>
          </div>
          <div className="previewScaleControl" aria-label="Preview display size">
            <span>表示サイズ</span>
            <button
              type="button"
              className="previewScaleButton"
              onClick={() =>
                setPreviewScalePercent((currentScale) =>
                  clampNumber(
                    currentScale - PREVIEW_SCALE_STEP,
                    PREVIEW_SCALE_MIN,
                    PREVIEW_SCALE_MAX,
                  ),
                )
              }
              disabled={previewScalePercent <= PREVIEW_SCALE_MIN}
              aria-label="Make preview smaller"
              title="Make preview smaller"
            >
              -
            </button>
            <output>{previewScalePercent}%</output>
            <button
              type="button"
              className="previewScaleButton"
              onClick={() =>
                setPreviewScalePercent((currentScale) =>
                  clampNumber(
                    currentScale + PREVIEW_SCALE_STEP,
                    PREVIEW_SCALE_MIN,
                    PREVIEW_SCALE_MAX,
                  ),
                )
              }
              disabled={previewScalePercent >= PREVIEW_SCALE_MAX}
              aria-label="Make preview larger"
              title="Make preview larger"
            >
              +
            </button>
          </div>
          <details
            className="controlsDetails controlsDetailsInline"
            style={{ display: 'none' }}
            aria-hidden="true"
          >
          <summary>詳細</summary>

          <p className="note">
            CSV columns: <code>order, page_no, step, side, body, image_1</code>
          </p>
          <p className="note">
            Apps Script route can also read the step cell background color, text color,
            and font family, plus shared SIDE blocks from a separate side sheet.
          </p>
          <p className="note">
            Use the <code>settings</code> sheet to set shared fonts like{' '}
            <code>step_font_family</code>, <code>body_font_family</code>, and{' '}
            <code>side_font_family</code>, plus layout values like{' '}
            <code>side_width</code>.
          </p>
          <p className="note">
            Put ids like <code>1,2,3</code> in <code>side</code>. When using Apps
            Script, those ids are resolved from the separate <code>side</code> sheet.
          </p>
          <p className="note">
            Use <code>direction</code> in the <code>side</code> sheet with{' '}
            <code>horizontal</code>, <code>vertical</code>, <code>upright</code>,{' '}
            <code>rotate_cw</code>, or <code>rotate_ccw</code>. The older{' '}
            <code>rotation</code> column is still supported.
          </p>
          <p className="note">
            Mix directions in one SIDE cell with <code>[h]...[/h]</code>,{' '}
            <code>[v]...[/v]</code>, <code>[u]...[/u]</code>,{' '}
            <code>[upright]...[/upright]</code>, <code>[tcy]...[/tcy]</code>,{' '}
            <code>[cw]...[/cw]</code>, or <code>[ccw]...[/ccw]</code>.
          </p>
          <p className="note">
            <code>image_1</code> accepts Google Drive share links and regular image URLs.
            Legacy column name <code>image</code> is also supported.
          </p>
          <p className="note">
            Put <code>{'{{image}}'}</code> inside <code>body</code> to place the image
            between text blocks. Use <code>{'{{image:2}}'}</code> for <code>image_2</code>.
            Use <code>{'{{images:1,2,3}}'}</code> to place multiple images in one row.
            Example: <code>{'{{images:1,2 width1=220px width2=120px align=center}}'}</code>
          </p>
          <p className="note">
            Use <code>{'[[hl]]text[[/hl]]'}</code> for a yellow highlight, or{' '}
            <code>{'[[hl:#ffd7a8]]text[[/hl]]'}</code> for a custom color.
          </p>
          <p className="note">
            Use <code>{'[[box]]text[[/box]]'}</code> for a red text box, or{' '}
            <code>{'[[box:#1f78c8]]text[[/box]]'}</code> for a custom color.
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
        </details>
        </div>
      </header>

      <section className="previewArea">
        {previewSpreads.map((spread) => (
          <section key={spread.key} className="sheetPreview">
            <div className="sheetMeta">
              <span>{spread.label}</span>
              <span>{spread.meta}</span>
            </div>
            <ResponsiveSheetCanvas displayScale={previewScalePercent / 100}>
                <PagePreview
                  page={spread.leftPage}
                  position="left"
                  layoutPosition={previewMode === 'book' ? 'right' : 'left'}
                  settings={result?.settings}
                  sideDefinitions={result?.sideDefinitions}
                />
                <PagePreview
                  page={spread.rightPage}
                  position="right"
                  layoutPosition={previewMode === 'book' ? 'left' : 'right'}
                  settings={result?.settings}
                  sideDefinitions={result?.sideDefinitions}
                />
            </ResponsiveSheetCanvas>
          </section>
        )) ?? null}
      </section>

      <nav className="scrollNav" aria-label="Page navigation">
        <button
          type="button"
          className="scrollNavButton"
          onClick={() => scrollToBoundary('top')}
          aria-label="ページトップに移動"
          title="ページトップに移動"
        >
          <ScrollNavIcon direction="top" />
        </button>
        <button
          type="button"
          className="scrollNavButton"
          onClick={() => scrollByPreviewPage('up')}
          aria-label="1ページ上に移動"
          title="1ページ上に移動"
        >
          <ScrollNavIcon direction="up" />
        </button>
        <button
          type="button"
          className="scrollNavButton"
          onClick={() => scrollByPreviewPage('down')}
          aria-label="1ページ下に移動"
          title="1ページ下に移動"
        >
          <ScrollNavIcon direction="down" />
        </button>
        <button
          type="button"
          className="scrollNavButton"
          onClick={() => scrollToBoundary('bottom')}
          aria-label="ページ最下部に移動"
          title="ページ最下部に移動"
        >
          <ScrollNavIcon direction="bottom" />
        </button>
      </nav>
    </main>
  )
}

export default App
