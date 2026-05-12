import type { PreviewColumn } from './types'

type CellValue = string | number | null
type Row = Record<string, CellValue>
export type ColumnFilters = Record<string, string[]>
const BLANK_FILTER_VALUE = '空白'

export function filterRows(rows: Row[], filters: ColumnFilters) {
  const activeFilters = Object.entries(filters)
    .map(([column, values]) => [
      column,
      values.map((value) => value.trim().toLowerCase()).filter(Boolean),
    ] as const)
    .filter(([, values]) => values.length > 0)

  return rows
    .map((row, sourceIndex) => ({ row, sourceIndex }))
    .filter(({ row }) =>
      activeFilters.every(([column, values]) =>
        values.includes(normalizedFilterText(row[column])),
      ),
    )
}

export function filterOptionsForColumn(rows: Row[], column: string) {
  const seen = new Set<string>()
  rows.forEach((row) => {
    seen.add(filterOptionText(row[column]))
  })
  return Array.from(seen).sort((left, right) => {
    if (left === BLANK_FILTER_VALUE) return 1
    if (right === BLANK_FILTER_VALUE) return -1
    return left.localeCompare(right, 'zh-Hans-CN', { numeric: true })
  })
}

export function nextColumnSelection(
  currentHeaders: string[],
  nextHeaders: string[],
  current: { groupColumns: string[]; sumColumns: string[] },
  guessed: { group: string[]; sum: string[] },
) {
  if (sameHeaders(currentHeaders, nextHeaders)) {
    return {
      groupColumns: current.groupColumns.filter((column) => nextHeaders.includes(column)),
      sumColumns: current.sumColumns.filter((column) => nextHeaders.includes(column)),
    }
  }

  return {
    groupColumns: guessed.group,
    sumColumns: guessed.sum,
  }
}

export function reorderColumns(columns: PreviewColumn[], fromKey: string, toKey: string) {
  if (fromKey === toKey) return columns

  const fromIndex = columns.findIndex((column) => column.key === fromKey)
  const toIndex = columns.findIndex((column) => column.key === toKey)
  if (fromIndex < 0 || toIndex < 0) return columns

  const next = [...columns]
  const [moved] = next.splice(fromIndex, 1)
  const adjustedToIndex = fromIndex < toIndex ? toIndex - 1 : toIndex
  next.splice(adjustedToIndex, 0, moved)

  return next.map((column, index) => ({
    ...column,
    index: index + 1,
    letter: excelColumnName(index + 1),
  }))
}

function sameHeaders(currentHeaders: string[], nextHeaders: string[]) {
  return currentHeaders.length === nextHeaders.length
    && currentHeaders.every((header, index) => header === nextHeaders[index])
}

function normalizedFilterText(value: CellValue | undefined) {
  return filterOptionText(value).trim().toLowerCase()
}

function filterOptionText(value: CellValue | undefined) {
  const text = cellToText(value).trim()
  return text || BLANK_FILTER_VALUE
}

function cellToText(value: CellValue | undefined) {
  if (value === null || value === undefined) return ''
  return String(value)
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
