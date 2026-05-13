import { beforeEach, describe, expect, it, vi } from 'vitest'

import {
  exportSummaryResults,
  fetchGuidePreference,
  fetchPreview,
  saveGuidePreference,
  uploadWorkbook,
} from './api'

const localWorkbook = vi.hoisted(() => ({
  uploadLocalWorkbook: vi.fn(),
  fetchLocalPreview: vi.fn(),
  exportLocalSummaryResults: vi.fn(),
  fetchLocalGuidePreference: vi.fn(),
  saveLocalGuidePreference: vi.fn(),
}))

vi.mock('./localWorkbook', () => localWorkbook)

describe('api local adapter', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    vi.stubGlobal('fetch', vi.fn(() => {
      throw new Error('fetch should not be called by the local adapter')
    }))
  })

  it('delegates workbook operations to the local workbook engine', async () => {
    const file = new File(['x'], 'sample.xlsx')
    localWorkbook.uploadLocalWorkbook.mockResolvedValue({ fileId: 'file-1' })
    localWorkbook.fetchLocalPreview.mockResolvedValue({ sheet_name: '明细' })
    localWorkbook.exportLocalSummaryResults.mockResolvedValue({ blob: new Blob(), fileName: 'out.xlsx' })

    await expect(uploadWorkbook(file)).resolves.toEqual({ fileId: 'file-1' })
    await expect(fetchPreview('file-1', '明细', 2)).resolves.toEqual({ sheet_name: '明细' })
    await expect(exportSummaryResults('file-1', [])).resolves.toMatchObject({ fileName: 'out.xlsx' })

    expect(localWorkbook.uploadLocalWorkbook).toHaveBeenCalledWith(file)
    expect(localWorkbook.fetchLocalPreview).toHaveBeenCalledWith('file-1', '明细', 2)
    expect(localWorkbook.exportLocalSummaryResults).toHaveBeenCalledWith('file-1', [])
    expect(fetch).not.toHaveBeenCalled()
  })

  it('delegates guide preference operations locally', async () => {
    localWorkbook.fetchLocalGuidePreference.mockResolvedValue({ guideHidden: false })
    localWorkbook.saveLocalGuidePreference.mockResolvedValue({ guideHidden: true })

    await expect(fetchGuidePreference()).resolves.toEqual({ guideHidden: false })
    await expect(saveGuidePreference(true)).resolves.toEqual({ guideHidden: true })

    expect(localWorkbook.fetchLocalGuidePreference).toHaveBeenCalled()
    expect(localWorkbook.saveLocalGuidePreference).toHaveBeenCalledWith(true)
  })
})
