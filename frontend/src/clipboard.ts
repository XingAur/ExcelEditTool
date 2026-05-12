import type { CellStyle, SummaryResult } from './types'

export function buildClipboardHtml(summary: SummaryResult): string {
  const columns = summary.columns
    .map((column) => `<col style="width:${column.width_px}px">`)
    .join('')
  const headerHeight = summary.row_heights['1'] ?? 26
  const dataHeight = summary.row_heights['2'] ?? 22

  const header = `<tr style="height:${headerHeight}px">${summary.columns
    .map((column) => {
      return `<th style="${styleToCss(summary.header_style)}">${escapeHtml(column.key)}</th>`
    })
    .join('')}</tr>`

  const body = summary.rows
    .map((row) => {
      const cells = summary.columns
        .map((column) => {
          const style = {
            ...summary.data_style,
            ...column.style,
          }
          return `<td style="${styleToCss(style)}">${escapeHtml(cellToText(row[column.key]))}</td>`
        })
        .join('')
      return `<tr style="height:${dataHeight}px">${cells}</tr>`
    })
    .join('')

  return [
    '<html><body>',
    '<table style="border-collapse:collapse;table-layout:fixed">',
    `<colgroup>${columns}</colgroup>`,
    `<thead>${header}</thead>`,
    `<tbody>${body}</tbody>`,
    '</table>',
    '</body></html>',
  ].join('')
}

export function buildClipboardText(summary: SummaryResult): string {
  const header = summary.columns.map((column) => column.key).join('\t')
  const rows = summary.rows.map((row) =>
    summary.columns.map((column) => cellToText(row[column.key])).join('\t'),
  )
  return [header, ...rows].join('\r\n')
}

export async function copySummaryToClipboard(summary: SummaryResult): Promise<void> {
  const html = buildClipboardHtml(summary)
  const text = buildClipboardText(summary)

  if ('ClipboardItem' in window && navigator.clipboard?.write) {
    const item = new ClipboardItem({
      'text/html': new Blob([html], { type: 'text/html' }),
      'text/plain': new Blob([text], { type: 'text/plain' }),
    })
    await navigator.clipboard.write([item])
    return
  }

  await navigator.clipboard.writeText(text)
}

export function styleToCss(style: CellStyle): string {
  const css: string[] = [
    'box-sizing:border-box',
    'padding:4px 8px',
    'white-space:nowrap',
    'overflow:hidden',
  ]

  if (style.background) css.push(`background-color:${style.background}`)
  if (style.color) css.push(`color:${style.color}`)
  if (style.bold) css.push('font-weight:700')
  if (style.horizontal) css.push(`text-align:${style.horizontal}`)
  if (style.vertical) css.push(`vertical-align:${style.vertical}`)
  if (style.font_name) css.push(`font-family:${style.font_name}`)
  if (style.font_size) css.push(`font-size:${style.font_size}pt`)

  const borderColor = style.border_color ?? '#D1D5DB'
  if (style.border_style) css.push(`border:1px solid ${borderColor}`)
  if (style.number_format && style.number_format !== 'General') {
    css.push(`mso-number-format:"${escapeCssValue(style.number_format)}"`)
  }

  return css.join(';')
}

function cellToText(value: string | number | null | undefined): string {
  if (value === null || value === undefined) return ''
  return String(value)
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;')
}

function escapeCssValue(value: string): string {
  return value.replaceAll('"', '\\"')
}
