import { describe, expect, it } from 'vitest'
import {
  filterRows,
  filterOptionsForColumn,
  nextColumnSelection,
  reorderColumns,
} from './tableTransforms'
import type { PreviewColumn } from './types'

const columns: PreviewColumn[] = [
  { key: '型号', index: 1, letter: 'A', width: 12, width_px: 89, style: {} },
  { key: '数量', index: 2, letter: 'B', width: 12, width_px: 89, style: {} },
  { key: '金额', index: 3, letter: 'C', width: 12, width_px: 89, style: {} },
]

describe('table transforms', () => {
  it('filters rows by multi-select dropdown values and keeps original row indexes', () => {
    const rows = [
      { 型号: 'A-100', 数量: 1, 金额: 10 },
      { 型号: 'B-200', 数量: 2, 金额: 20 },
      { 型号: 'A-300', 数量: 3, 金额: 30 },
    ]

    const result = filterRows(rows, { 型号: ['A-100', 'A-300'], 数量: ['3'] })

    expect(result).toEqual([
      { row: { 型号: 'A-300', 数量: 3, 金额: 30 }, sourceIndex: 2 },
    ])
  })

  it('builds dropdown options from a column without duplicates', () => {
    const rows = [
      { 型号: 'A-100', 数量: 1, 金额: 10 },
      { 型号: 'A-100', 数量: 2, 金额: 20 },
      { 型号: '', 数量: 3, 金额: 30 },
      { 型号: null, 数量: 4, 金额: 40 },
    ]

    expect(filterOptionsForColumn(rows, '型号')).toEqual(['A-100', '空白'])
  })

  it('does not filter a column when its selected values are empty', () => {
    const rows = [
      { 型号: 'A-100', 数量: 1, 金额: 10 },
      { 型号: 'B-200', 数量: 2, 金额: 20 },
    ]

    expect(filterRows(rows, { 型号: [] }).map(({ sourceIndex }) => sourceIndex)).toEqual([0, 1])
  })

  it('moves a dragged column before the target column and refreshes letters', () => {
    const result = reorderColumns(columns, '金额', '型号')

    expect(result.map((column) => column.key)).toEqual(['金额', '型号', '数量'])
    expect(result.map((column) => [column.index, column.letter])).toEqual([
      [1, 'A'],
      [2, 'B'],
      [3, 'C'],
    ])
  })

  it('keeps selected columns when switching sheets with the same headers', () => {
    const current = {
      groupColumns: ['型号'],
      sumColumns: ['金额'],
    }

    expect(
      nextColumnSelection(
        ['型号', '数量', '金额'],
        ['型号', '数量', '金额'],
        current,
        { group: ['型号'], sum: ['数量', '金额'] },
      ),
    ).toEqual(current)
  })

  it('uses guessed columns when switching sheets with different headers', () => {
    const guessed = { group: ['产品'], sum: ['金额'] }

    expect(
      nextColumnSelection(
        ['型号', '数量'],
        ['产品', '金额'],
        { groupColumns: ['型号'], sumColumns: ['数量'] },
        guessed,
      ),
    ).toEqual({
      groupColumns: ['产品'],
      sumColumns: ['金额'],
    })
  })
})
