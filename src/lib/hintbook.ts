import { parseCsv } from './csv'

export type SheetRecord = {
  rowNumber: number
  order: number
  pageNo: string
  step: string
  side: string
  body: string
  image: string
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
}

const fieldAliases = {
  order: ['order', 'sort', 'sort_order'],
  pageNo: ['page_no', 'page', 'pageno', 'page_number'],
  step: ['step', 'stage', 'section'],
  side: ['side', 'label', 'side_label'],
  body: ['body', 'text', 'content'],
  image: ['image', 'image_key', 'image_url'],
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

function parseOrder(value: string) {
  const parsed = Number(value)
  return Number.isInteger(parsed) && parsed > 0 ? parsed : null
}

function orderChunks<T>(items: T[], size: number) {
  const chunks: T[][] = []

  for (let index = 0; index < items.length; index += size) {
    chunks.push(items.slice(index, index + size))
  }

  return chunks
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
  const warnings: string[] = []

  for (const field of requiredFields) {
    if (!fieldMap.includes(field)) {
      warnings.push(`Missing column: ${field}`)
    }
  }

  const orderMap = new Map<number, SheetRecord>()

  dataRows.forEach((row, index) => {
    const values = fieldMap.reduce<Record<string, string>>(
      (accumulator, field, headerIndex) => {
        if (!field) return accumulator
        accumulator[field] = row[headerIndex]?.trim() ?? ''
        return accumulator
      },
      {},
    )

    const hasAnyValue = Object.values(values).some(Boolean)
    if (!hasAnyValue) {
      return
    }

    const order = parseOrder(values.order ?? '')
    if (order === null) {
      return
    }

    if (orderMap.has(order)) {
      warnings.push(`Duplicate order ${order}: row ${index + 2} was ignored.`)
      return
    }

    orderMap.set(order, {
      rowNumber: index + 2,
      order,
      pageNo: values.pageNo ?? '',
      step: values.step ?? '',
      side: values.side ?? '',
      body: values.body ?? '',
      image: values.image ?? '',
    })
  })

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
  }
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

export const sheetColumnGuide = [
  ['order', 'Required. Real sequence number used for print layout. Blank rows are ignored.'],
  ['page_no', 'Display page number. This does not control sorting.'],
  ['step', 'Top label of the page.'],
  ['side', 'Side label text shown temporarily on the page.'],
  ['body', 'Free text body.'],
  ['image', 'Image URL or key. Placeholder only for now.'],
] as const
