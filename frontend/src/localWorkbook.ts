import ExcelJS from 'exceljs'

import type {
  CellStyle,
  GuidePreference,
  SheetPreview,
  SummaryResult,
  UploadResponse,
} from './types'

type StoredWorkbook = {
  fileName: string
  workbook: ExcelJS.Workbook
}

const workbookStore = new Map<string, StoredWorkbook>()
const GUIDE_STORAGE_KEY = 'excel-edit-tool-guide-hidden'
let fallbackGuideHidden = false

export async function uploadLocalWorkbook(file: File): Promise<UploadResponse> {
  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    throw new Error('请上传 .xlsx 文件')
  }

  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(await file.arrayBuffer())
  const fileId = crypto.randomUUID()
  workbookStore.set(fileId, { fileName: file.name, workbook })

  const preview = createPreview(file.name, workbook, workbook.worksheets[0]?.name, undefined)
  return {
    fileId,
    fileName: file.name,
    sheets: workbook.worksheets.map((sheet) => sheet.name),
    preview,
  }
}

export async function fetchLocalPreview(
  fileId: string,
  sheetName: string,
  headerRow: number,
): Promise<SheetPreview> {
  const stored = getStoredWorkbook(fileId)
  return createPreview(stored.fileName, stored.workbook, sheetName, headerRow)
}

export async function exportLocalSummaryResults(
  fileId: string,
  results: SummaryResult[],
): Promise<{ blob: Blob; fileName: string }> {
  const stored = getStoredWorkbook(fileId)
  const workbook = new ExcelJS.Workbook()

  results.forEach((result) => {
    const sheet = workbook.addWorksheet(safeSheetName(result.sheet_name))
    result.columns.forEach((column, index) => {
      const cell = sheet.getCell(1, index + 1)
      cell.value = column.key
      applyCellStyle(cell, { ...result.header_style, ...column.style })
      sheet.getColumn(index + 1).width = column.width || 12
    })

    result.rows.forEach((row, rowIndex) => {
      result.columns.forEach((column, columnIndex) => {
        const cell = sheet.getCell(rowIndex + 2, columnIndex + 1)
        cell.value = row[column.key] ?? ''
        applyCellStyle(cell, result.data_style)
      })
    })
  })

  const buffer = await workbook.xlsx.writeBuffer()
  return {
    blob: new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    fileName: summaryFileName(stored.fileName),
  }
}

export async function fetchLocalGuidePreference(): Promise<GuidePreference> {
  const storage = getLocalStorage()
  return { guideHidden: storage ? storage.getItem(GUIDE_STORAGE_KEY) === '1' : fallbackGuideHidden }
}

export async function saveLocalGuidePreference(guideHidden: boolean): Promise<GuidePreference> {
  const storage = getLocalStorage()
  if (storage) {
    if (guideHidden) storage.setItem(GUIDE_STORAGE_KEY, '1')
    else storage.removeItem(GUIDE_STORAGE_KEY)
  } else {
    fallbackGuideHidden = guideHidden
  }
  return { guideHidden }
}

function getStoredWorkbook(fileId: string): StoredWorkbook {
  const stored = workbookStore.get(fileId)
  if (!stored) {
    throw new Error('文件不存在或已过期')
  }
  return stored
}

function createPreview(
  fileName: string,
  workbook: ExcelJS.Workbook,
  sheetName?: string,
  requestedHeaderRow?: number,
): SheetPreview {
  const sheet = sheetName ? workbook.getWorksheet(sheetName) : workbook.worksheets[0]
  if (!sheet) {
    throw new Error('工作表不存在')
  }

  const headerRow = requestedHeaderRow ?? detectHeaderRow(sheet)
  const headerValues = rowValues(sheet.getRow(headerRow))
  const headers = headerValues.map((value, index) => normalizeHeader(value, index + 1))
  const rows = collectRows(sheet, headerRow, headers)
  const columns = headers.map((header, index) => {
    const column = sheet.getColumn(index + 1)
    return {
      key: header,
      index: index + 1,
      letter: excelColumnName(index + 1),
      width: Number(column.width ?? 12),
      width_px: Math.round(Number(column.width ?? 12) * 7 + 5),
      style: cellStyle(sheet.getCell(headerRow, index + 1)),
    }
  })

  return {
    file_name: fileName,
    sheets: workbook.worksheets.map((item) => item.name),
    sheet_name: sheet.name,
    header_row: headerRow,
    headers,
    rows,
    columns,
    row_heights: collectRowHeights(sheet),
    header_style: cellStyle(sheet.getRow(headerRow).getCell(1)),
    data_style: cellStyle(sheet.getRow(headerRow + 1).getCell(1)),
  }
}

function detectHeaderRow(sheet: ExcelJS.Worksheet): number {
  let bestRow = 1
  let bestScore = -1
  sheet.eachRow((row, rowNumber) => {
    const values = rowValues(row)
      .map((value) => String(value ?? '').trim())
      .filter(Boolean)
    const score = new Set(values).size * 10 + values.length
    if (score > bestScore) {
      bestScore = score
      bestRow = rowNumber
    }
  })
  return bestRow
}

function collectRows(
  sheet: ExcelJS.Worksheet,
  headerRow: number,
  headers: string[],
): Record<string, string | number | null>[] {
  const rows: Record<string, string | number | null>[] = []
  for (let rowNumber = headerRow + 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber)
    const values = headers.map((header, index) => [header, normalizeCellValue(row.getCell(index + 1).value)] as const)
    if (!values.some(([, value]) => value !== null && value !== '')) continue
    rows.push(Object.fromEntries(values))
  }
  return rows
}

function collectRowHeights(sheet: ExcelJS.Worksheet): Record<string, number> {
  const rowHeights: Record<string, number> = {}
  sheet.eachRow((row, rowNumber) => {
    if (row.height) rowHeights[String(rowNumber)] = row.height
  })
  return rowHeights
}

function rowValues(row: ExcelJS.Row): Array<string | number | null> {
  const values: Array<string | number | null> = []
  const max = Math.max(row.cellCount, row.actualCellCount)
  for (let index = 1; index <= max; index += 1) {
    values.push(normalizeCellValue(row.getCell(index).value))
  }
  return values
}

function normalizeHeader(value: string | number | null, index: number): string {
  const text = String(value ?? '').trim()
  return text || `列${index}`
}

function normalizeCellValue(value: ExcelJS.CellValue): string | number | null {
  if (value === null || value === undefined) return null
  if (typeof value === 'string' || typeof value === 'number') return value
  if (value instanceof Date) return value.toISOString()
  if (typeof value === 'object') {
    if ('result' in value) return normalizeCellValue(value.result as ExcelJS.CellValue)
    if ('text' in value) return String(value.text)
    if ('richText' in value) return value.richText.map((item) => item.text).join('')
    if ('formula' in value) return String(value.result ?? '')
  }
  return String(value)
}

function cellStyle(cell: ExcelJS.Cell): CellStyle {
  return {
    background: argbToHex(cell.fill?.type === 'pattern' ? cell.fill.fgColor?.argb : undefined),
    color: argbToHex(cell.font?.color?.argb),
    bold: cell.font?.bold ?? false,
    horizontal: cell.alignment?.horizontal ?? null,
    vertical: cell.alignment?.vertical ?? null,
    number_format: cell.numFmt ?? null,
    border_color: borderColor(cell),
    border_style: cell.border ? 'thin' : null,
    font_name: cell.font?.name ?? null,
    font_size: cell.font?.size ?? null,
  }
}

function applyCellStyle(cell: ExcelJS.Cell, style: CellStyle) {
  if (style.bold || style.color || style.font_name || style.font_size) {
    cell.font = {
      bold: style.bold,
      color: style.color ? { argb: hexToArgb(style.color) } : undefined,
      name: style.font_name ?? undefined,
      size: style.font_size ?? undefined,
    }
  }
  if (style.background) {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: hexToArgb(style.background) },
    }
  }
  if (style.horizontal || style.vertical) {
    cell.alignment = {
      horizontal: style.horizontal as ExcelJS.Alignment['horizontal'],
      vertical: style.vertical as ExcelJS.Alignment['vertical'],
    }
  }
}

function borderColor(cell: ExcelJS.Cell): string | null {
  const border = cell.border?.top ?? cell.border?.right ?? cell.border?.bottom ?? cell.border?.left
  return argbToHex(border?.color?.argb)
}

function argbToHex(argb?: string): string | null {
  if (!argb) return null
  const hex = argb.length === 8 ? argb.slice(2) : argb
  return `#${hex}`
}

function hexToArgb(hex: string): string {
  return `FF${hex.replace('#', '').toUpperCase()}`
}

function summaryFileName(fileName: string): string {
  const stem = fileName.replace(/\.[^.]+$/, '')
  return `${stem}_汇总.xlsx`
}

function safeSheetName(name: string): string {
  return name.replace(/[\\/*?:[\]]/g, '').slice(0, 31) || '汇总结果'
}

function excelColumnName(index: number) {
  let name = ''
  let current = index
  while (current > 0) {
    const remainder = (current - 1) % 26
    name = String.fromCharCode(65 + remainder) + name
    current = Math.floor((current - 1) / 26)
  }
  return name
}

function getLocalStorage(): Storage | null {
  try {
    return globalThis.localStorage ?? null
  } catch {
    return null
  }
}
