import { describe, expect, it } from 'vitest'
import { buildEditedSummary, updateRowCell } from './summary'
import type { SheetPreview, SummaryConfig } from './types'

const preview: SheetPreview = {
  file_name: 'sample.xlsx',
  sheets: ['上海', '安徽'],
  sheet_name: '上海',
  header_row: 1,
  headers: ['型号', '单位', '数量', '金额'],
  rows: [
    { 型号: 'A', 单位: '台', 数量: 1, 金额: 10 },
    { 型号: 'A', 单位: '台', 数量: 2, 金额: 20 },
  ],
  columns: [
    { key: '型号', index: 1, letter: 'A', width: 12, width_px: 89, style: {} },
    { key: '单位', index: 2, letter: 'B', width: 12, width_px: 89, style: {} },
    { key: '数量', index: 3, letter: 'C', width: 12, width_px: 89, style: {} },
    { key: '金额', index: 4, letter: 'D', width: 12, width_px: 89, style: {} },
  ],
  row_heights: { 1: 24, 2: 22 },
  header_style: {},
  data_style: {},
}

const config: SummaryConfig = {
  sheet_name: '上海',
  header_row: 1,
  group_columns: ['型号', '单位'],
  sum_columns: ['数量', '金额'],
}

describe('edited summary', () => {
  it('summarizes edited source rows instead of the original preview rows', () => {
    const editedRows = updateRowCell(preview.rows, 1, '数量', '5')
    const result = buildEditedSummary(preview, editedRows, config)

    expect(result.sheet_name).toBe('上海')
    expect(result.rows).toEqual([{ 型号: 'A', 单位: '台', 数量: 6, 金额: 30 }])
  })

  it('keeps manually edited summary rows available for export', () => {
    const result = buildEditedSummary(preview, preview.rows, config)
    const editedSummaryRows = updateRowCell(result.rows, 0, '金额', '99.5')

    expect(editedSummaryRows[0].金额).toBe(99.5)
  })

  it('uses original Excel row numbers when summarizing filtered rows', () => {
    const filteredRows = [{ 型号: 'A', 单位: '台', 数量: '坏数据', 金额: 20 }]
    const result = buildEditedSummary(preview, filteredRows, config, [1])

    expect(result.warnings).toEqual(['第 3 行的 数量 不是数字，已按 0 处理'])
  })
})
