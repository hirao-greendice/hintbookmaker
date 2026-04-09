import { parseCsv } from './csv'

export type StepStyle = {
  backgroundColor?: string
  textColor?: string
  fontFamily?: string
  fontSize?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strikethrough?: boolean
}

export type RichTextRun = {
  text: string
  textColor?: string
  fontFamily?: string
  fontSize?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strikethrough?: boolean
}

export type RenderSettings = {
  stepFontFamily?: string
  bodyFontFamily?: string
  sideFontFamily?: string
  pageNoFontFamily?: string
  stepFontScale?: number
  bodyFontScale?: number
}

export type ImageSources = Record<string, string>

export type SideBlockDefinition = {
  id: string
  text: string
  textRuns?: RichTextRun[]
  height?: number
  backgroundColor?: string
  textColor?: string
  fontFamily?: string
  fontSize?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strikethrough?: boolean
}

export type SideBlockDefinitions = Record<string, SideBlockDefinition>

export type SheetRecord = {
  rowNumber: number
  order: number
  pageNo: string
  step: string
  side: string
  body: string
  image: string
  imagePosition?: string
  imageWidth?: string
  imageHeight?: string
  imageAlign?: string
  imageFit?: string
  imageSources?: ImageSources
  stepStyle?: StepStyle
  bodyStyle?: StepStyle
  stepRuns?: RichTextRun[]
  bodyRuns?: RichTextRun[]
}

export type FormattedPage = SheetRecord

export type PrintSpread = {
  sheetNumber: number
  spreadNumber: number
  leftPage: FormattedPage | null
  rightPage: FormattedPage | null
}

export type FormatResult = {
  rawRows: string[][]
  records: SheetRecord[]
  pages: FormattedPage[]
  spreads: PrintSpread[]
  warnings: string[]
  settings?: RenderSettings
  sideDefinitions?: SideBlockDefinitions
}

type RowSeed = {
  rowNumber: number
  order: string
  pageNo: string
  step: string
  side: string
  body: string
  image: string
  imagePosition?: string
  imageWidth?: string
  imageHeight?: string
  imageAlign?: string
  imageFit?: string
  imageSources?: ImageSources
  stepStyle?: StepStyle
  bodyStyle?: StepStyle
  stepRuns?: RichTextRun[]
  bodyRuns?: RichTextRun[]
}

type AppsScriptRow = Record<string, unknown> & {
  stepStyle?: unknown
  step_style?: unknown
}

type AppsScriptPayload = {
  rows?: unknown
  settings?: unknown
  sideDefinitions?: unknown
}

const fieldAliases = {
  order: ['order', 'sort', 'sort_order'],
  pageNo: ['page_no', 'page', 'pageno', 'page_number'],
  step: ['step', 'stage', 'section'],
  side: ['side', 'label', 'side_label'],
  body: ['body', 'text', 'content'],
  image: ['image', 'image_key', 'image_url'],
  imagePosition: ['image_position', 'image_pos', 'image_place'],
  imageWidth: ['image_width', 'image_w'],
  imageHeight: ['image_height', 'image_h'],
  imageAlign: ['image_align', 'image_alignment'],
  imageFit: ['image_fit', 'image_object_fit'],
} as const

const requiredFields: Array<keyof typeof fieldAliases> = [
  'order',
  'pageNo',
  'step',
  'side',
  'body',
  'image',
]

function normalizeHeader(value: string) {
  return value
    .trim()
    .toLowerCase()
    .replace(/[\s-]+/g, '_')
}

function aliasToField(header: string) {
  const normalized = normalizeHeader(header)
  const entries = Object.entries(fieldAliases) as Array<
    [keyof typeof fieldAliases, readonly string[]]
  >

  const match = entries.find(([, aliases]) =>
    aliases.some((alias) => normalizeHeader(alias) === normalized),
  )

  return match?.[0] ?? null
}

function getImageFieldIndex(header: string) {
  const normalized = normalizeHeader(header)
  if (normalized === 'image') {
    return '1'
  }

  const match = normalized.match(/^image_?(\d+)$/)
  return match?.[1] ?? null
}

function parseOrder(value: string) {
  const parsed = Number(value)
  return Number.isInteger(parsed) && parsed > 0 ? parsed : null
}

function collectImageSources(
  entries: Array<[string, unknown]>,
): ImageSources | undefined {
  const imageSources: ImageSources = {}

  for (const [key, value] of entries) {
    const index = getImageFieldIndex(key)
    if (!index) {
      continue
    }

    const text = toStringValue(value).trim()
    if (!text) {
      continue
    }

    imageSources[index] = text
  }

  return Object.keys(imageSources).length > 0 ? imageSources : undefined
}

function toStringValue(value: unknown) {
  if (value === null || value === undefined) return ''
  return String(value)
}

function sanitizeColor(value: unknown) {
  const text = toStringValue(value).trim()
  if (!text) return undefined
  return text
}

function sanitizeFontFamily(value: unknown) {
  const text = toStringValue(value).trim()
  if (!text) return undefined
  return text
}

function sanitizeFontSize(value: unknown) {
  const parsed = Number(value)
  if (!Number.isFinite(parsed) || parsed <= 0) return undefined
  return parsed
}

function sanitizeBoolean(value: unknown) {
  if (value === true || value === false) return value
  const text = toStringValue(value).trim().toLowerCase()
  if (!text) return undefined
  if (text === 'true') return true
  if (text === 'false') return false
  return undefined
}

function parseStepStyle(value: unknown): StepStyle | undefined {
  if (!value || typeof value !== 'object') {
    return undefined
  }

  const record = value as Record<string, unknown>
  const backgroundColor =
    sanitizeColor(record.backgroundColor) ??
    sanitizeColor(record.background_color) ??
    sanitizeColor(record.bgColor) ??
    sanitizeColor(record.bg_color)

  const textColor =
    sanitizeColor(record.textColor) ??
    sanitizeColor(record.text_color) ??
    sanitizeColor(record.fontColor) ??
    sanitizeColor(record.font_color)

  const fontFamily =
    sanitizeFontFamily(record.fontFamily) ??
    sanitizeFontFamily(record.font_family)
  const fontSize =
    sanitizeFontSize(record.fontSize) ??
    sanitizeFontSize(record.font_size)
  const bold = sanitizeBoolean(record.bold)
  const italic = sanitizeBoolean(record.italic)
  const underline = sanitizeBoolean(record.underline)
  const strikethrough =
    sanitizeBoolean(record.strikethrough) ??
    sanitizeBoolean(record.strikeThrough) ??
    sanitizeBoolean(record.strike_through)

  if (
    !backgroundColor &&
    !textColor &&
    !fontFamily &&
    !fontSize &&
    bold === undefined &&
    italic === undefined &&
    underline === undefined &&
    strikethrough === undefined
  ) {
    return undefined
  }

  return {
    backgroundColor,
    textColor,
    fontFamily,
    fontSize,
    bold,
    italic,
    underline,
    strikethrough,
  }
}

function parseRichTextRuns(value: unknown): RichTextRun[] | undefined {
  if (!Array.isArray(value)) {
    return undefined
  }

  const runs: RichTextRun[] = []

  for (const entry of value) {
    if (!entry || typeof entry !== 'object') {
      continue
    }

    const record = entry as Record<string, unknown>
    const text = toStringValue(record.text)
    if (text === '') {
      continue
    }

    runs.push({
      text,
      textColor:
        sanitizeColor(record.textColor) ?? sanitizeColor(record.text_color),
      fontFamily:
        sanitizeFontFamily(record.fontFamily) ??
        sanitizeFontFamily(record.font_family),
      fontSize:
        sanitizeFontSize(record.fontSize) ??
        sanitizeFontSize(record.font_size),
      bold: sanitizeBoolean(record.bold),
      italic: sanitizeBoolean(record.italic),
      underline: sanitizeBoolean(record.underline),
      strikethrough:
        sanitizeBoolean(record.strikethrough) ??
        sanitizeBoolean(record.strikeThrough) ??
        sanitizeBoolean(record.strike_through),
    })
  }

  return runs.length > 0 ? runs : undefined
}

function parseSideBlockDefinitions(value: unknown): SideBlockDefinitions | undefined {
  if (!value || typeof value !== 'object') {
    return undefined
  }

  const entries = Object.entries(value as Record<string, unknown>)
  const definitions: SideBlockDefinitions = {}

  for (const [id, entry] of entries) {
    if (!entry || typeof entry !== 'object') {
      continue
    }

    const record = entry as Record<string, unknown>
    const text = toStringValue(record.text).trim()
    const textRuns = parseRichTextRuns(record.textRuns ?? record.text_runs)
    const height = sanitizeFontSize(record.height)
    const backgroundColor = sanitizeColor(record.backgroundColor)
    const textColor = sanitizeColor(record.textColor)
    const fontFamily = sanitizeFontFamily(record.fontFamily)
    const fontSize = sanitizeFontSize(record.fontSize)
    const bold = sanitizeBoolean(record.bold)
    const italic = sanitizeBoolean(record.italic)
    const underline = sanitizeBoolean(record.underline)
    const strikethrough =
      sanitizeBoolean(record.strikethrough) ??
      sanitizeBoolean(record.strikeThrough) ??
      sanitizeBoolean(record.strike_through)

    if (
      !text &&
      !textRuns &&
      height === undefined &&
      !backgroundColor &&
      !textColor &&
      !fontFamily &&
      fontSize === undefined &&
      bold === undefined &&
      italic === undefined &&
      underline === undefined &&
      strikethrough === undefined
    ) {
      continue
    }

    definitions[id] = {
      id,
      text,
      textRuns,
      height,
      backgroundColor,
      textColor,
      fontFamily,
      fontSize,
      bold,
      italic,
      underline,
      strikethrough,
    }
  }

  return Object.keys(definitions).length > 0 ? definitions : undefined
}

function orderChunks<T>(items: T[], size: number) {
  const chunks: T[][] = []

  for (let index = 0; index < items.length; index += size) {
    chunks.push(items.slice(index, index + size))
  }

  return chunks
}

function parseRenderSettings(value: unknown): RenderSettings | undefined {
  if (!value || typeof value !== 'object') {
    return undefined
  }

  const record = value as Record<string, unknown>
  const settings: RenderSettings = {
    stepFontFamily:
      sanitizeFontFamily(record.stepFontFamily) ??
      sanitizeFontFamily(record.step_font_family),
    bodyFontFamily:
      sanitizeFontFamily(record.bodyFontFamily) ??
      sanitizeFontFamily(record.body_font_family),
    sideFontFamily:
      sanitizeFontFamily(record.sideFontFamily) ??
      sanitizeFontFamily(record.side_font_family),
    pageNoFontFamily:
      sanitizeFontFamily(record.pageNoFontFamily) ??
      sanitizeFontFamily(record.page_no_font_family),
    stepFontScale:
      sanitizeFontSize(record.stepFontScale) ??
      sanitizeFontSize(record.step_font_scale),
    bodyFontScale:
      sanitizeFontSize(record.bodyFontScale) ??
      sanitizeFontSize(record.body_font_scale),
  }

  if (
    !settings.stepFontFamily &&
    !settings.bodyFontFamily &&
    !settings.sideFontFamily &&
    !settings.pageNoFontFamily &&
    !settings.stepFontScale &&
    !settings.bodyFontScale
  ) {
    return undefined
  }

  return settings
}

function buildPrintSpreads(pages: FormattedPage[]): PrintSpread[] {
  const chunks = orderChunks(pages, 4)

  return chunks.flatMap((chunk, chunkIndex) => {
    const first = chunk[0] ?? null
    const second = chunk[1] ?? null
    const third = chunk[2] ?? null
    const fourth = chunk[3] ?? null
    const baseSheetNumber = chunkIndex * 2 + 1

    return [
      {
        sheetNumber: baseSheetNumber,
        spreadNumber: chunkIndex + 1,
        leftPage: first,
        rightPage: fourth,
      },
      {
        sheetNumber: baseSheetNumber + 1,
        spreadNumber: chunkIndex + 1,
        leftPage: third,
        rightPage: second,
      },
    ]
  })
}

function buildResultFromSeeds(
  rawRows: string[][],
  seeds: RowSeed[],
  settings?: RenderSettings,
  sideDefinitions?: SideBlockDefinitions,
): FormatResult {
  const warnings: string[] = []
  const orderMap = new Map<number, SheetRecord>()

  for (const seed of seeds) {
    const hasAnyValue = [
      seed.order,
      seed.pageNo,
      seed.step,
      seed.side,
      seed.body,
      seed.image,
    ].some((value) => value.trim() !== '')

    if (!hasAnyValue) {
      continue
    }

    const order = parseOrder(seed.order)
    if (order === null) {
      continue
    }

    if (orderMap.has(order)) {
      warnings.push(`Duplicate order ${order}: row ${seed.rowNumber} was ignored.`)
      continue
    }

    orderMap.set(order, {
      rowNumber: seed.rowNumber,
      order,
      pageNo: seed.pageNo,
      step: seed.step,
      side: seed.side,
      body: seed.body,
      image: seed.image,
      imagePosition: seed.imagePosition,
      imageWidth: seed.imageWidth,
      imageHeight: seed.imageHeight,
      imageAlign: seed.imageAlign,
      imageFit: seed.imageFit,
      imageSources: seed.imageSources,
      stepStyle: seed.stepStyle,
      bodyStyle: seed.bodyStyle,
      stepRuns: seed.stepRuns,
      bodyRuns: seed.bodyRuns,
    })
  }

  const records = Array.from(orderMap.values()).sort(
    (first, second) => first.order - second.order,
  )

  for (let index = 0; index < records.length; index += 1) {
    const expectedOrder = index + 1
    if (records[index]?.order !== expectedOrder) {
      warnings.push(`Missing order ${expectedOrder}.`)
    }
  }

  const pages = records
  const spreads = buildPrintSpreads(pages)

  return {
    rawRows,
    records,
    pages,
    spreads,
    warnings,
    settings,
    sideDefinitions,
  }
}

export function buildGoogleSheetCsvUrl(input: string) {
  const trimmed = input.trim()
  if (!trimmed) {
    throw new Error('Spreadsheet URL or ID is required.')
  }

  if (trimmed.includes('format=csv') || trimmed.includes('output=csv')) {
    return trimmed
  }

  if (/^[a-zA-Z0-9-_]+$/.test(trimmed) && !trimmed.startsWith('http')) {
    return `https://docs.google.com/spreadsheets/d/${trimmed}/export?format=csv`
  }

  let url: URL
  try {
    url = new URL(trimmed)
  } catch {
    throw new Error('Could not parse spreadsheet URL or ID.')
  }

  const match = url.pathname.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/)
  if (!match) {
    throw new Error('This is not a Google Sheets URL.')
  }

  const spreadsheetId = match[1]
  const gid = url.searchParams.get('gid')
  const params = new URLSearchParams({ format: 'csv' })

  if (gid) {
    params.set('gid', gid)
  }

  return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?${params.toString()}`
}

export function formatHintBookFromCsv(csvText: string): FormatResult {
  const rawRows = parseCsv(csvText)
  if (rawRows.length === 0) {
    throw new Error('CSV is empty.')
  }

  const [headerRow, ...dataRows] = rawRows
  const fieldMap = headerRow.map((header) => aliasToField(header))
  const missing = requiredFields.filter((field) => !fieldMap.includes(field))

  if (missing.length > 0) {
    throw new Error(`Missing columns: ${missing.join(', ')}`)
  }

  const seeds = dataRows.map((row, index) => {
    const rawEntries = headerRow.map((header, headerIndex) => [
      header,
      row[headerIndex] ?? '',
    ]) as Array<[string, unknown]>
    const values = fieldMap.reduce<Record<string, string>>(
      (accumulator, field, headerIndex) => {
        if (!field) return accumulator
        accumulator[field] = row[headerIndex]?.trim() ?? ''
        return accumulator
      },
      {},
    )

    return {
      rowNumber: index + 2,
      order: values.order ?? '',
      pageNo: values.pageNo ?? '',
      step: values.step ?? '',
      side: values.side ?? '',
      body: values.body ?? '',
      image: values.image ?? '',
      imagePosition: values.imagePosition ?? '',
      imageWidth: values.imageWidth ?? '',
      imageHeight: values.imageHeight ?? '',
      imageAlign: values.imageAlign ?? '',
      imageFit: values.imageFit ?? '',
      imageSources: collectImageSources(rawEntries),
      stepStyle: undefined,
      bodyStyle: undefined,
      stepRuns: undefined,
      bodyRuns: undefined,
    } satisfies RowSeed
  })

  return buildResultFromSeeds(rawRows, seeds)
}

export function formatHintBookFromAppsScript(payload: unknown): FormatResult {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Apps Script response is not an object.')
  }

  const container = payload as AppsScriptPayload
  const rows = Array.isArray(container.rows) ? container.rows : null
  if (!rows) {
    throw new Error('Apps Script response must have a rows array.')
  }
  const settings = parseRenderSettings(container.settings)
  const sideDefinitions = parseSideBlockDefinitions(container.sideDefinitions)

  const seeds = rows.map((entry, index) => {
    const source = entry && typeof entry === 'object' ? (entry as AppsScriptRow) : {}
    const sourceEntries = Object.entries(source)
    const values: Partial<Record<keyof typeof fieldAliases, string>> = {}

    for (const [key, value] of sourceEntries) {
      const field = aliasToField(key)
      if (field) {
        values[field] = toStringValue(value).trim()
      }
    }

    return {
      rowNumber: index + 2,
      order: values.order ?? '',
      pageNo: values.pageNo ?? '',
      step: values.step ?? '',
      side: values.side ?? '',
      body: values.body ?? '',
      image: values.image ?? '',
      imagePosition: values.imagePosition ?? '',
      imageWidth: values.imageWidth ?? '',
      imageHeight: values.imageHeight ?? '',
      imageAlign: values.imageAlign ?? '',
      imageFit: values.imageFit ?? '',
      imageSources: collectImageSources(sourceEntries),
      stepStyle: parseStepStyle(source.stepStyle ?? source.step_style),
      bodyStyle: parseStepStyle(source.bodyStyle ?? source.body_style),
      stepRuns: parseRichTextRuns(source.stepRuns ?? source.step_runs),
      bodyRuns: parseRichTextRuns(source.bodyRuns ?? source.body_runs),
    } satisfies RowSeed
  })

  return buildResultFromSeeds([], seeds, settings, sideDefinitions)
}

export const defaultSheetSource =
  'https://docs.google.com/spreadsheets/d/1d4XuVJPSy579inDl082_Qr84CHU5rFpcnx0kSwiblg4/edit?usp=sharing'

export const sampleCsv = `order,page_no,step,side,body,image
1,1,全体目次,1st,"1st STEP
赤のページへ",
2,2,1st-1,1st,"謎ID:001のヒント
イラストのルールを見るページです。",
3,3,1st-1,2nd,"謎ID:002のヒント
補足のテキストが入ります。",
4,4,1st-1,1st,"謎ID:003のヒント
ここも仮の本文です。",`

export const sampleAppsScriptResponse = {
  settings: {
    step_font_family: 'MS Mincho',
    body_font_family: 'Yu Gothic',
    side_font_family: 'MS Mincho',
    page_no_font_family: 'Arial',
    step_font_scale: 4,
    body_font_scale: 2.5,
  },
  sideDefinitions: {
    '1': {
      id: '1',
      text: '転換\n1',
      height: 12,
      backgroundColor: '#d94b67',
      textColor: '#ffffff',
      fontFamily: 'Noto Serif JP',
      bold: true,
    },
    '2': {
      id: '2',
      text: '転換\n2',
      height: 12,
      backgroundColor: '#1f78c8',
      textColor: '#ffffff',
      fontFamily: 'Noto Serif JP',
      bold: true,
    },
    '3': {
      id: '3',
      text: '転換\n3',
      height: 12,
      backgroundColor: '#54a33f',
      textColor: '#ffffff',
      fontFamily: 'Noto Serif JP',
      bold: true,
    },
  },
  rows: [
    {
      order: 1,
      page_no: '1',
      step: '全体目次',
      side: '1,2,3',
      body: '1st STEP\n赤のページへ',
      image: '',
      stepStyle: {
        backgroundColor: '#5c5c5c',
        textColor: '#ffffff',
        fontFamily: 'Noto Serif JP',
      },
      stepRuns: [
        {
          text: '全体目次',
          textColor: '#ffffff',
          fontFamily: 'Noto Serif JP',
          bold: true,
        },
      ],
      bodyRuns: [
        {
          text: '1st STEP\n',
          bold: true,
        },
        {
          text: '赤のページへ',
        },
      ],
    },
    {
      order: 2,
      page_no: '2',
      step: '1st-1',
      side: '2,3',
      body: '謎ID:001のヒント\nイラストのルールを見るページです。',
      image: '',
      stepStyle: {
        backgroundColor: '#d94b67',
        textColor: '#ffffff',
        fontFamily: 'Noto Serif JP',
      },
      bodyRuns: [
        {
          text: '謎ID:001',
          bold: true,
          textColor: '#d94b67',
        },
        {
          text: 'のヒント\nイラストのルールを見るページです。',
        },
      ],
    },
  ],
}

export const sheetColumnGuide = [
  ['order', 'Required. Real sequence number used for print layout. Blank rows are ignored.'],
  ['page_no', 'Display page number. This does not control sorting.'],
  ['step', 'Top label of the page.'],
  ['side', 'Comma-separated SIDE block ids such as 1,2,3. These ids are resolved from the separate side sheet when using Apps Script.'],
  ['body', 'Free text body. Use {{image}}, {{image:2}}, {{image:3}} for single images, or {{images:1,2,3}} for one horizontal row of multiple images.'],
  ['image', 'Primary image URL, Google Drive share link, or linked cell.'],
  ['image_2', 'Optional second image source. Also supports image_3, image_4, and so on.'],
  ['image_position', 'Optional fallback. top or bottom. Used only when BODY does not contain {{image}}.'],
  ['image_width', 'Optional. CSS width such as 160px or 60%. Numbers are treated as px.'],
  ['image_height', 'Optional. CSS height such as 120px or 40%. Numbers are treated as px.'],
  ['image_align', 'Optional. left, center, or right.'],
  ['image_fit', 'Optional. contain or cover.'],
] as const
