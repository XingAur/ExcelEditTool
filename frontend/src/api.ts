import {
  exportLocalSummaryResults,
  fetchLocalGuidePreference,
  fetchLocalPreview,
  saveLocalGuidePreference,
  uploadLocalWorkbook,
} from './localWorkbook'
import type {
  GuidePreference,
  SheetPreview,
  SummaryConfig,
  SummaryResult,
  UploadResponse,
} from './types'

export async function uploadWorkbook(file: File): Promise<UploadResponse> {
  return uploadLocalWorkbook(file)
}

export async function fetchPreview(
  fileId: string,
  sheetName: string,
  headerRow: number,
): Promise<SheetPreview> {
  return fetchLocalPreview(fileId, sheetName, headerRow)
}

export async function createSummary(
  fileId: string,
  config: SummaryConfig,
): Promise<SummaryResult> {
  void fileId
  void config
  throw new Error('汇总结果在本地界面中生成')
}

export async function exportSummary(
  fileId: string,
  config: SummaryConfig,
): Promise<{ blob: Blob; fileName: string }> {
  void fileId
  void config
  throw new Error('请先生成汇总结果后再导出')
}

export async function exportSummaryResults(
  fileId: string,
  results: SummaryResult[],
): Promise<{ blob: Blob; fileName: string }> {
  return exportLocalSummaryResults(fileId, results)
}

export async function fetchGuidePreference(): Promise<GuidePreference> {
  return fetchLocalGuidePreference()
}

export async function saveGuidePreference(guideHidden: boolean): Promise<GuidePreference> {
  return saveLocalGuidePreference(guideHidden)
}
