export type CellStyle = {
  background?: string | null
  color?: string | null
  bold?: boolean
  horizontal?: string | null
  vertical?: string | null
  number_format?: string | null
  border_color?: string | null
  border_style?: string | null
  font_name?: string | null
  font_size?: number | null
}

export type PreviewColumn = {
  key: string
  index: number
  letter: string
  width: number
  width_px: number
  style: CellStyle
}

export type SheetPreview = {
  file_name: string
  sheets: string[]
  sheet_name: string
  header_row: number
  headers: string[]
  rows: Record<string, string | number | null>[]
  columns: PreviewColumn[]
  row_heights: Record<string, number>
  header_style: CellStyle
  data_style: CellStyle
}

export type SummaryConfig = {
  sheet_name: string
  header_row: number
  group_columns: string[]
  sum_columns: string[]
}

export type SummaryResult = {
  sheet_name: string
  header_row: number
  columns: PreviewColumn[]
  rows: Record<string, string | number | null>[]
  warnings: string[]
  row_heights: Record<string, number>
  header_style: CellStyle
  data_style: CellStyle
}

export type SummaryExportRequest = {
  results: SummaryResult[]
}

export type UploadResponse = {
  fileId: string
  fileName: string
  sheets: string[]
  preview: SheetPreview
}

export type GuidePreference = {
  guideHidden: boolean
}
