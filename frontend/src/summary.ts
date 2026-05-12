import type { PreviewColumn, SheetPreview, SummaryConfig, SummaryResult } from './types'

type CellValue = string | number | null
type Row = Record<string, CellValue>

export function buildEditedSummary(
  preview: SheetPreview,
  rows: Row[],
  config: SummaryConfig,
  sourceRowIndexes?: number[],
): SummaryResult {
  const outputHeaders = [...config.group_columns, ...config.sum_columns]
  const columns = outputHeaders
    .map((header, index) => {
      const source = preview.columns.find((column) => column.key === header)
      return {
        ...(source ?? fallbackColumn(header, index + 1)),
        index: index + 1,
        letter: excelColumnName(index + 1),
      }
    })

  const warnings: string[] = []
  const groups = new Map<string, Row>()

  rows.forEach((row, index) => {
    if (!Object.values(row).some((value) => value !== null && value !== '')) return
    const keyValues = config.group_columns.map((column) => normalizeKey(row[column]))
    const key = JSON.stringify(keyValues)

    if (!groups.has(key)) {
      const initial: Row = {}
      config.group_columns.forEach((column) => {
        initial[column] = row[column] ?? ''
      })
      config.sum_columns.forEach((column) => {
        initial[column] = 0
      })
      groups.set(key, initial)
    }

    const target = groups.get(key)!
    config.sum_columns.forEach((column) => {
      const current = toNumber(target[column])
      const next = parseNumeric(row[column])
      if (next === null) {
        const sourceIndex = sourceRowIndexes?.[index] ?? index
        warnings.push(`第 ${sourceIndex + preview.header_row + 1} 行的 ${column} 不是数字，已按 0 处理`)
        return
      }
      target[column] = roundMoney(current + next)
    })
  })

  return {
    sheet_name: preview.sheet_name,
    header_row: 1,
    columns,
    rows: Array.from(groups.values()),
    warnings,
    row_heights: preview.row_heights,
    header_style: preview.header_style,
    data_style: preview.data_style,
  }
}

export function updateRowCell(rows: Row[], rowIndex: number, column: string, value: string): Row[] {
  return rows.map((row, index) =>
    index === rowIndex ? { ...row, [column]: coerceCellValue(value) } : row,
  )
}

function fallbackColumn(header: string, index: number): PreviewColumn {
  return {
    key: header,
    index,
    letter: excelColumnName(index),
    width: 12,
    width_px: 89,
    style: {},
  }
}

function normalizeKey(value: CellValue | undefined): string {
  if (value === null || value === undefined) return ''
  return String(value).trim()
}

function parseNumeric(value: CellValue | undefined): number | null {
  if (value === null || value === undefined || value === '') return 0
  const normalized = String(value).replaceAll(',', '').replaceAll('￥', '').replaceAll('¥', '').trim()
  if (normalized === '') return 0
  const parsed = Number(normalized)
  return Number.isFinite(parsed) ? parsed : null
}

function toNumber(value: CellValue | undefined): number {
  const parsed = parseNumeric(value)
  return parsed ?? 0
}

function roundMoney(value: number): number {
  return Math.round((value + Number.EPSILON) * 1000000) / 1000000
}

function coerceCellValue(value: string): CellValue {
  const trimmed = value.trim()
  if (trimmed === '') return ''
  const parsed = Number(trimmed.replaceAll(',', ''))
  return Number.isFinite(parsed) && /^-?\d+(\.\d+)?$/.test(trimmed.replaceAll(',', ''))
    ? parsed
    : value
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
