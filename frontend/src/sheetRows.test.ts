import { describe, expect, it } from 'vitest'
import { compactRowNumberWidth, rowsForPreview, sheetRowsKey } from './sheetRows'
import type { SheetPreview } from './types'

function makePreview(headerRow: number, rows: Array<Record<string, string | number | null>>): SheetPreview {
  return {
    file_name: 'sample.xlsx',
    sheets: ['明细'],
    sheet_name: '明细',
    header_row: headerRow,
    headers: ['型号', '数量'],
    rows,
    columns: [
      { key: '型号', index: 1, letter: 'A', width: 12, width_px: 89, style: {} },
      { key: '数量', index: 2, letter: 'B', width: 12, width_px: 89, style: {} },
    ],
    row_heights: { 1: 24, 2: 22 },
    header_style: {},
    data_style: {},
  }
}

describe('sheet row cache helpers', () => {
  it('keeps edited rows separate when the header row changes', () => {
    const firstHeaderPreview = makePreview(1, [{ 型号: 'A', 数量: 1 }])
    const secondHeaderPreview = makePreview(2, [{ 型号: 'B', 数量: 2 }])
    const store = {
      [sheetRowsKey('明细', 1)]: [{ 型号: 'edited', 数量: 9 }],
    }

    expect(rowsForPreview(store, firstHeaderPreview)).toEqual([{ 型号: 'edited', 数量: 9 }])
    expect(rowsForPreview(store, secondHeaderPreview)).toEqual([{ 型号: 'B', 数量: 2 }])
  })

  it('keeps the row number column compact', () => {
    expect(compactRowNumberWidth(9)).toBe(18)
    expect(compactRowNumberWidth(10)).toBe(20)
    expect(compactRowNumberWidth(9999)).toBe(32)
  })
})
