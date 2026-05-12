import type { SheetPreview } from './types'

type GridRow = Record<string, string | number | null>
export type RowStore = Record<string, GridRow[]>

export function sheetRowsKey(sheetName: string, headerRow: number) {
  return `${sheetName}::header-${headerRow}`
}

export function rowsForPreview(store: RowStore, preview: SheetPreview) {
  return store[sheetRowsKey(preview.sheet_name, preview.header_row)] ?? preview.rows
}

export function compactRowNumberWidth(lastRowNumber: number) {
  return Math.min(34, Math.max(18, String(lastRowNumber).length * 6 + 8))
}
