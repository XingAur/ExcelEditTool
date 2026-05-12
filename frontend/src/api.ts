import type {
  GuidePreference,
  SheetPreview,
  SummaryConfig,
  SummaryResult,
  UploadResponse,
} from './types'

const API_BASE = import.meta.env.VITE_API_BASE ?? ''

export async function uploadWorkbook(file: File): Promise<UploadResponse> {
  const body = new FormData()
  body.append('file', file)
  const response = await fetch(`${API_BASE}/api/workbooks`, {
    method: 'POST',
    body,
  })
  return readJson(response)
}

export async function fetchPreview(
  fileId: string,
  sheetName: string,
  headerRow: number,
): Promise<SheetPreview> {
  const params = new URLSearchParams({
    sheetName,
    headerRow: String(headerRow),
  })
  const response = await fetch(`${API_BASE}/api/workbooks/${fileId}/preview?${params}`)
  return readJson(response)
}

export async function createSummary(
  fileId: string,
  config: SummaryConfig,
): Promise<SummaryResult> {
  const response = await fetch(`${API_BASE}/api/workbooks/${fileId}/summary`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(config),
  })
  return readJson(response)
}

export async function exportSummary(
  fileId: string,
  config: SummaryConfig,
): Promise<{ blob: Blob; fileName: string }> {
  const response = await fetch(`${API_BASE}/api/workbooks/${fileId}/export`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(config),
  })

  if (!response.ok) {
    throw new Error(await readError(response))
  }

  return {
    blob: await response.blob(),
    fileName: parseDownloadFileName(response.headers.get('content-disposition')),
  }
}

export async function exportSummaryResults(
  fileId: string,
  results: SummaryResult[],
): Promise<{ blob: Blob; fileName: string }> {
  const response = await fetch(`${API_BASE}/api/workbooks/${fileId}/export-results`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ results }),
  })

  if (!response.ok) {
    throw new Error(await readError(response))
  }

  return {
    blob: await response.blob(),
    fileName: parseDownloadFileName(response.headers.get('content-disposition')),
  }
}

export async function fetchGuidePreference(): Promise<GuidePreference> {
  const response = await fetch(`${API_BASE}/api/preferences/guide`)
  return readJson(response)
}

export async function saveGuidePreference(guideHidden: boolean): Promise<GuidePreference> {
  const response = await fetch(`${API_BASE}/api/preferences/guide`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ guideHidden }),
  })
  return readJson(response)
}

async function readJson<T>(response: Response): Promise<T> {
  if (!response.ok) {
    throw new Error(await readError(response))
  }
  return response.json()
}

async function readError(response: Response): Promise<string> {
  try {
    const payload = await response.json()
    return payload.detail ?? '请求失败'
  } catch {
    return '请求失败'
  }
}

function parseDownloadFileName(disposition: string | null): string {
  if (!disposition) return '汇总结果.xlsx'
  const encoded = disposition.match(/filename\*=UTF-8''([^;]+)/)
  if (encoded?.[1]) return decodeURIComponent(encoded[1])
  const plain = disposition.match(/filename="?([^";]+)"?/)
  return plain?.[1] ?? '汇总结果.xlsx'
}
