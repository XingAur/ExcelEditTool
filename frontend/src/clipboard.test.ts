import { describe, expect, it } from 'vitest'
import { buildClipboardHtml, buildClipboardText } from './clipboard'
import type { SummaryResult } from './types'

const summary: SummaryResult = {
  sheet_name: '汇总结果',
  header_row: 1,
  columns: [
    {
      key: '规格型号',
      index: 1,
      letter: 'A',
      width: 28,
      width_px: 201,
      style: {
        background: '#E2F0D9',
        border_color: '#D1D5DB',
        border_style: 'thin',
      },
    },
    {
      key: '金额',
      index: 2,
      letter: 'B',
      width: 12,
      width_px: 89,
      style: {
        number_format: '#,##0.00',
        border_color: '#D1D5DB',
        border_style: 'thin',
      },
    },
  ],
  rows: [{ 规格型号: 'US880', 金额: 2772 }],
  warnings: [],
  row_heights: { 1: 26, 2: 22 },
  header_style: {
    background: '#217346',
    color: '#FFFFFF',
    bold: true,
    horizontal: 'center',
    vertical: 'center',
    border_color: '#D1D5DB',
    border_style: 'thin',
  },
  data_style: {
    background: '#E2F0D9',
    vertical: 'center',
    border_color: '#D1D5DB',
    border_style: 'thin',
  },
}

describe('clipboard serialization', () => {
  it('builds HTML with column widths, row heights, and cell styles', () => {
    const html = buildClipboardHtml(summary)

    expect(html).toContain('<col style="width:201px"')
    expect(html).toContain('height:26px')
    expect(html).toContain('background-color:#217346')
    expect(html).toContain('font-weight:700')
    expect(html).toContain('US880')
    expect(html).toContain('2772')
  })

  it('builds tab separated plain text for Excel fallback paste', () => {
    expect(buildClipboardText(summary)).toBe('规格型号\t金额\r\nUS880\t2772')
  })
})
