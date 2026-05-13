import ExcelJS from 'exceljs'
import { describe, expect, it } from 'vitest'

import {
  exportLocalSummaryResults,
  fetchLocalPreview,
  saveLocalGuidePreference,
  fetchLocalGuidePreference,
  uploadLocalWorkbook,
} from './localWorkbook'
import type { SummaryResult } from './types'

async function buildWorkbookFile() {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('明细')
  sheet.getRow(1).values = ['说明', '说明', '说明']
  sheet.getRow(2).values = ['规格型号', '单价', '金额']
  sheet.getRow(3).values = ['US880', 12, 24]
  sheet.getRow(4).values = ['US880', 12, 36]
  sheet.getRow(2).font = { bold: true, color: { argb: 'FFFFFFFF' } }
  sheet.getRow(2).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF4472C4' },
  }
  sheet.getColumn(1).width = 18

  const buffer = await workbook.xlsx.writeBuffer()
  return new File([buffer], '销售明细.xlsx', {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
}

describe('local workbook engine', () => {
  it('uploads and previews a workbook without a server', async () => {
    const upload = await uploadLocalWorkbook(await buildWorkbookFile())

    expect(upload.fileName).toBe('销售明细.xlsx')
    expect(upload.sheets).toEqual(['明细'])
    expect(upload.preview.sheet_name).toBe('明细')
    expect(upload.preview.header_row).toBe(2)
    expect(upload.preview.headers).toEqual(['规格型号', '单价', '金额'])
    expect(upload.preview.rows).toEqual([
      { 规格型号: 'US880', 单价: 12, 金额: 24 },
      { 规格型号: 'US880', 单价: 12, 金额: 36 },
    ])
    expect(upload.preview.columns[0]).toMatchObject({
      key: '规格型号',
      index: 1,
      letter: 'A',
      width: 18,
    })

    const preview = await fetchLocalPreview(upload.fileId, '明细', 2)
    expect(preview.rows).toHaveLength(2)
  })

  it('exports edited summary results as an xlsx workbook', async () => {
    const upload = await uploadLocalWorkbook(await buildWorkbookFile())
    const result: SummaryResult = {
      sheet_name: '明细',
      header_row: 1,
      columns: upload.preview.columns.filter((column) => ['规格型号', '金额'].includes(column.key)),
      rows: [{ 规格型号: 'US880', 金额: 60 }],
      warnings: [],
      row_heights: {},
      header_style: upload.preview.header_style,
      data_style: upload.preview.data_style,
    }

    const { blob, fileName } = await exportLocalSummaryResults(upload.fileId, [result])
    const exported = new ExcelJS.Workbook()
    await exported.xlsx.load(await blob.arrayBuffer())

    expect(fileName).toBe('销售明细_汇总.xlsx')
    expect(exported.worksheets.map((sheet) => sheet.name)).toEqual(['明细'])
    expect(exported.getWorksheet('明细')?.getCell('A1').value).toBe('规格型号')
    expect(exported.getWorksheet('明细')?.getCell('B2').value).toBe(60)
  })

  it('stores guide preference locally', async () => {
    await saveLocalGuidePreference(true)
    await expect(fetchLocalGuidePreference()).resolves.toEqual({ guideHidden: true })
  })
})
