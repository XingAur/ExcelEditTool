import type {
  CSSProperties,
  ChangeEvent,
  DragEvent,
  KeyboardEvent as ReactKeyboardEvent,
  PointerEvent as ReactPointerEvent,
  ReactNode,
} from 'react'
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import {
  Button,
  Checkbox,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Input,
  Spinner,
  Text,
} from '@fluentui/react-components'
import {
  ArrowDownload24Regular,
  ArrowUpload24Regular,
  Calculator24Regular,
  CheckmarkCircle20Filled,
  Copy24Regular,
  DocumentTable24Regular,
  Filter16Regular,
} from '@fluentui/react-icons'
import { copySummaryToClipboard } from './clipboard'
import {
  exportSummaryResults,
  fetchGuidePreference,
  fetchPreview,
  saveGuidePreference,
  uploadWorkbook,
} from './api'
import { buildEditedSummary, updateRowCell } from './summary'
import { compactRowNumberWidth, rowsForPreview, sheetRowsKey } from './sheetRows'
import {
  filterOptionsForColumn,
  filterRows,
  nextColumnSelection,
  reorderColumns,
  type ColumnFilters,
} from './tableTransforms'
import type { CellStyle, PreviewColumn, SheetPreview, SummaryConfig, SummaryResult } from './types'
import './styles.css'

type CellValue = string | number | null
type GridRow = Record<string, CellValue>
type RowStore = Record<string, GridRow[]>
type SummaryStore = Record<string, SummaryResult>
type FilterStore = Record<string, ColumnFilters>
const GUIDE_STORAGE_KEY = 'excel-edit-tool-guide-hidden'
const EMPTY_FILTERS: ColumnFilters = {}
const FILTER_MENU_WIDTH = 224
const FILTER_MENU_MAX_HEIGHT = 320
const FILTER_MENU_MARGIN = 8
const FILTER_MENU_GAP = 6

function App() {
  const fileInputRef = useRef<HTMLInputElement>(null)
  const [fileId, setFileId] = useState('')
  const [fileName, setFileName] = useState('')
  const [sheets, setSheets] = useState<string[]>([])
  const [selectedSheet, setSelectedSheet] = useState('')
  const [selectedSummarySheet, setSelectedSummarySheet] = useState('')
  const [headerRow, setHeaderRow] = useState(1)
  const [preview, setPreview] = useState<SheetPreview | null>(null)
  const [sourceRowsBySheet, setSourceRowsBySheet] = useState<RowStore>({})
  const [sourceFiltersBySheet, setSourceFiltersBySheet] = useState<FilterStore>({})
  const [summaryBySheet, setSummaryBySheet] = useState<SummaryStore>({})
  const [groupColumns, setGroupColumns] = useState<string[]>([])
  const [sumColumns, setSumColumns] = useState<string[]>([])
  const [loading, setLoading] = useState('')
  const [error, setError] = useState('')
  const [copied, setCopied] = useState(false)
  const [guideOpen, setGuideOpen] = useState(false)
  const [dontShowGuide, setDontShowGuide] = useState(false)

  const activeSummary = selectedSummarySheet ? summaryBySheet[selectedSummarySheet] : null
  const generatedSummarySheets = sheets.filter((sheet) => Boolean(summaryBySheet[sheet]))
  const sourceFilterKey = preview ? sheetRowsKey(preview.sheet_name, preview.header_row) : ''
  const sourceFilters = sourceFilterKey ? sourceFiltersBySheet[sourceFilterKey] ?? EMPTY_FILTERS : EMPTY_FILTERS
  const filteredSourceRows = useMemo(
    () => (preview ? filterRows(preview.rows, sourceFilters) : []),
    [preview, sourceFilters],
  )
  const visibleSourceRows = useMemo(
    () => filteredSourceRows.map(({ row }) => row),
    [filteredSourceRows],
  )
  const visibleSourceRowIndexes = useMemo(
    () => filteredSourceRows.map(({ sourceIndex }) => sourceIndex),
    [filteredSourceRows],
  )
  const activeFilterCount = Object.values(sourceFilters).filter((values) => values.length > 0).length

  const summaryConfig = useMemo<SummaryConfig | null>(() => {
    if (!selectedSheet) return null
    return {
      sheet_name: selectedSheet,
      header_row: preview?.header_row ?? headerRow,
      group_columns: groupColumns,
      sum_columns: sumColumns,
    }
  }, [groupColumns, headerRow, preview?.header_row, selectedSheet, sumColumns])

  useEffect(() => {
    let cancelled = false
    fetchGuidePreference()
      .then((preference) => {
        if (!cancelled) setGuideOpen(!preference.guideHidden)
      })
      .catch(() => {
        if (!cancelled) setGuideOpen(shouldShowGuide())
      })
    return () => {
      cancelled = true
    }
  }, [])

  async function handleUpload(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    if (!file) return
    setError('')
    setCopied(false)
    setLoading('upload')
    try {
      const payload = await uploadWorkbook(file)
      setFileId(payload.fileId)
      setFileName(payload.fileName)
      setSheets(payload.sheets)
      setSelectedSheet(payload.preview.sheet_name)
      setSelectedSummarySheet('')
      setHeaderRow(payload.preview.header_row)
      setPreview(payload.preview)
      setSourceRowsBySheet({
        [sheetRowsKey(payload.preview.sheet_name, payload.preview.header_row)]: payload.preview.rows,
      })
      setSourceFiltersBySheet({})
      setSummaryBySheet({})
      const guessed = guessColumns(payload.preview.headers)
      setGroupColumns(guessed.group)
      setSumColumns(guessed.sum)
    } catch (err) {
      setError(errorMessage(err))
    } finally {
      setLoading('')
    }
  }

  async function loadPreview(sheetName = selectedSheet, row = headerRow) {
    if (!fileId || !sheetName) return
    setError('')
    setCopied(false)
    setLoading('preview')
    try {
      const nextPreview = await fetchPreview(fileId, sheetName, row)
      const rows = rowsForPreview(sourceRowsBySheet, nextPreview)
      const cacheKey = sheetRowsKey(nextPreview.sheet_name, nextPreview.header_row)
      setPreview({ ...nextPreview, rows })
      setSourceRowsBySheet((current) => ({
        ...current,
        [cacheKey]: current[cacheKey] ?? nextPreview.rows,
      }))
      setSelectedSheet(nextPreview.sheet_name)
      setHeaderRow(nextPreview.header_row)
      const guessed = guessColumns(nextPreview.headers)
      const nextSelection = nextColumnSelection(
        preview?.headers ?? [],
        nextPreview.headers,
        { groupColumns, sumColumns },
        guessed,
      )
      setGroupColumns(nextSelection.groupColumns)
      setSumColumns(nextSelection.sumColumns)
    } catch (err) {
      setError(errorMessage(err))
    } finally {
      setLoading('')
    }
  }

  function handleSourceFilterChange(column: string, values: string[]) {
    if (!preview) return
    const key = sheetRowsKey(preview.sheet_name, preview.header_row)
    setSourceFiltersBySheet((current) => {
      const nextFilters = { ...(current[key] ?? {}) }
      if (values.length > 0) nextFilters[column] = values
      else delete nextFilters[column]
      return { ...current, [key]: nextFilters }
    })
  }

  function handleSourceCellChange(rowIndex: number, column: string, value: string) {
    if (!preview || !selectedSheet) return
    const nextRows = updateRowCell(preview.rows, rowIndex, column, value)
    setPreview({ ...preview, rows: nextRows })
    setSourceRowsBySheet((current) => ({
      ...current,
      [sheetRowsKey(selectedSheet, preview.header_row)]: nextRows,
    }))
  }

  function handleSummaryColumnReorder(fromKey: string, toKey: string) {
    if (!activeSummary || !selectedSummarySheet) return
    const nextSummary = {
      ...activeSummary,
      columns: reorderColumns(activeSummary.columns, fromKey, toKey),
    }
    setSummaryBySheet((current) => ({ ...current, [selectedSummarySheet]: nextSummary }))
  }

  function handleSummaryCellChange(rowIndex: number, column: string, value: string) {
    if (!activeSummary || !selectedSummarySheet) return
    const nextSummary = {
      ...activeSummary,
      rows: updateRowCell(activeSummary.rows, rowIndex, column, value),
    }
    setSummaryBySheet((current) => ({ ...current, [selectedSummarySheet]: nextSummary }))
  }

  function handleCreateSummary() {
    if (!preview || !summaryConfig) return
    if (summaryConfig.group_columns.length === 0 || summaryConfig.sum_columns.length === 0) {
      setError('请选择至少一个分组列和一个求和列')
      return
    }
    setError('')
    setCopied(false)
    setLoading('summary')
    try {
      const result = buildEditedSummary(
        preview,
        visibleSourceRows,
        summaryConfig,
        visibleSourceRowIndexes,
      )
      setSummaryBySheet((current) => ({ ...current, [selectedSheet]: result }))
      setSelectedSummarySheet(selectedSheet)
    } catch (err) {
      setError(errorMessage(err))
    } finally {
      setLoading('')
    }
  }

  async function handleCopy() {
    if (!activeSummary) return
    setError('')
    try {
      await copySummaryToClipboard(activeSummary)
      setCopied(true)
      window.setTimeout(() => setCopied(false), 2000)
    } catch (err) {
      setError(`复制失败：${errorMessage(err)}`)
    }
  }

  async function handleExport() {
    if (!fileId) return
    const results = sheets.map((sheet) => summaryBySheet[sheet]).filter(Boolean)
    if (results.length === 0) {
      setError('请先生成至少一个 sheet 的汇总结果')
      return
    }
    setError('')
    setLoading('export')
    try {
      const { blob, fileName: downloadName } = await exportSummaryResults(fileId, results)
      await saveBlobWithPicker(blob, downloadName)
    } catch (err) {
      if (err instanceof DOMException && err.name === 'AbortError') return
      setError(errorMessage(err))
    } finally {
      setLoading('')
    }
  }

  function toggleGroup(column: string, checked: boolean) {
    setGroupColumns((current) =>
      checked ? [...new Set([...current, column])] : current.filter((item) => item !== column),
    )
    if (checked) setSumColumns((current) => current.filter((item) => item !== column))
  }

  function toggleSum(column: string, checked: boolean) {
    setSumColumns((current) =>
      checked ? [...new Set([...current, column])] : current.filter((item) => item !== column),
    )
    if (checked) setGroupColumns((current) => current.filter((item) => item !== column))
  }

  function closeGuide() {
    if (dontShowGuide) {
      rememberGuidePreference(true)
      void saveGuidePreference(true).catch(() => undefined)
    }
    setGuideOpen(false)
  }

  return (
    <main className="shell">
      <header className="commandBar">
        <div className="brand">
          <div className="brandIcon">
            <DocumentTable24Regular />
          </div>
          <div>
            <h1>ExcelEditTool</h1>
            <p>本地处理，多 sheet 预览，自定义汇总，双击编辑后复制与导出</p>
          </div>
        </div>

        <div className="commandActions">
          <input
            ref={fileInputRef}
            className="hiddenInput"
            type="file"
            accept=".xlsx"
            onChange={handleUpload}
          />
          <Button
            appearance="primary"
            icon={<ArrowUpload24Regular />}
            onClick={() => fileInputRef.current?.click()}
          >
            选择 Excel
          </Button>
        </div>
      </header>

      {error && <div className="message errorMessage">{error}</div>}
      {loading && (
        <div className="message statusMessage">
          <Spinner size="tiny" />
          <span>{loadingLabel(loading)}</span>
        </div>
      )}

      <section className="workspace">
        <section className="gridArea">
          <ResizableSplit
            className="previewSplit"
            ariaLabel="调整原始数据和汇总结果比例"
            firstClassName="previewPane"
            secondClassName="previewPane"
            first={
              <>
                <div className="tableHeader">
                  <div>
                    <Text weight="semibold">原始数据预览</Text>
                    <p>
                      {preview
                        ? `${fileName} · ${preview.sheet_name} · 总行数 ${preview.rows.length}${activeFilterCount ? ` · 过滤后 ${visibleSourceRows.length} 行` : ''}`
                        : '等待上传 Excel'}
                    </p>
                  </div>
                  <HeaderRowControl
                    headerRow={headerRow}
                    disabled={!fileId || loading === 'preview'}
                    onHeaderRowChange={setHeaderRow}
                    onApply={() => loadPreview()}
                  />
                </div>

                <ExcelWorkbookFrame
                  columns={preview?.columns ?? []}
                  rows={visibleSourceRows}
                  rowIndexes={visibleSourceRowIndexes}
                  firstRowNumber={preview?.header_row ?? 1}
                  headerStyle={preview?.header_style}
                  dataStyle={preview?.data_style}
                  rowHeights={preview?.row_heights}
                  emptyText="选择 Excel 后在这里预览原始数据"
                  headerFilters={sourceFilters}
                  filterOptionRows={preview?.rows ?? []}
                  onHeaderFilterChange={handleSourceFilterChange}
                  onCellChange={handleSourceCellChange}
                  footer={
                    <SheetTabs
                      sheets={sheets}
                      selectedSheet={selectedSheet}
                      disabled={!fileId || loading === 'preview'}
                      onSelect={(sheet) => loadPreview(sheet, headerRow)}
                    />
                  }
                />
              </>
            }
            second={
              <>
                <div className="tableHeader resultTitle">
                  <div>
                    <Text weight="semibold">汇总结果预览</Text>
                    <p>{activeSummary ? `${activeSummary.sheet_name} · 结果总行数 ${activeSummary.rows.length}` : '生成后显示新的结果页'}</p>
                  </div>
                  <ResultActions
                    canUseResult={Boolean(activeSummary)}
                    copied={copied}
                    exporting={loading === 'export'}
                    onCopy={handleCopy}
                    onExport={handleExport}
                  />
                </div>

                <ExcelWorkbookFrame
                  columns={activeSummary?.columns ?? []}
                  rows={activeSummary?.rows ?? []}
                  firstRowNumber={activeSummary?.header_row ?? 1}
                  headerStyle={activeSummary?.header_style}
                  dataStyle={activeSummary?.data_style}
                  rowHeights={activeSummary?.row_heights}
                  emptyText="汇总结果会显示在这里"
                  draggableColumns={Boolean(activeSummary)}
                  onCellChange={handleSummaryCellChange}
                  onColumnReorder={handleSummaryColumnReorder}
                  footer={
                    <SheetTabs
                      sheets={generatedSummarySheets}
                      selectedSheet={selectedSummarySheet}
                      disabled={false}
                      emptyText="生成汇总后显示 sheet"
                      onSelect={setSelectedSummarySheet}
                    />
                  }
                />

                {activeSummary?.warnings.length ? (
                  <div className="warningList">
                    {activeSummary.warnings.map((warning) => (
                      <div key={warning}>{warning}</div>
                    ))}
                  </div>
                ) : null}
              </>
            }
          />
        </section>

        <aside className="configPanel">
          <div className="panelTitle">
            <Calculator24Regular />
            <span>汇总设置</span>
          </div>

          <div className="configBody">
            <ResizableSplit
              className="configSplit"
              ariaLabel="调整分组列和求和列比例"
              firstClassName="configBlock"
              secondClassName="configBlock"
              first={
                <>
                  <h2>分组列</h2>
                  <p>相同组合会合并为一行</p>
                  <ColumnChecks
                    headers={preview?.headers ?? []}
                    selected={groupColumns}
                    disabled={sumColumns}
                    onToggle={toggleGroup}
                  />
                </>
              }
              second={
                <>
                  <h2>求和列</h2>
                  <p>数量、金额等数字列会求和</p>
                  <ColumnChecks
                    headers={preview?.headers ?? []}
                    selected={sumColumns}
                    disabled={groupColumns}
                    onToggle={toggleSum}
                  />
                </>
              }
            />
          </div>

          <div className="configFooter">
            <Button
              className="generateButton"
              appearance="primary"
              size="large"
              icon={<Calculator24Regular />}
              disabled={!fileId || loading === 'summary'}
              onClick={handleCreateSummary}
            >
              生成汇总
            </Button>
          </div>
        </aside>
      </section>

      <GuideDialog
        open={guideOpen}
        dontShowAgain={dontShowGuide}
        onDontShowAgainChange={setDontShowGuide}
        onClose={closeGuide}
      />
    </main>
  )
}

export function ResizableSplit({
  first,
  second,
  ariaLabel,
  className,
  firstClassName,
  secondClassName,
  initialRatio = 0.5,
  minRatio = 0.24,
  maxRatio = 0.76,
}: {
  first: ReactNode
  second: ReactNode
  ariaLabel: string
  className?: string
  firstClassName?: string
  secondClassName?: string
  initialRatio?: number
  minRatio?: number
  maxRatio?: number
}) {
  const containerRef = useRef<HTMLDivElement>(null)
  const [ratio, setRatio] = useState(() => clampRatio(initialRatio, minRatio, maxRatio))
  const firstPercent = formatRatioPercent(ratio)
  const secondPercent = formatRatioPercent(1 - ratio)

  function setRatioFromClientY(clientY: number) {
    const rect = containerRef.current?.getBoundingClientRect()
    if (!rect || rect.height <= 0) return
    setRatio(clampRatio((clientY - rect.top) / rect.height, minRatio, maxRatio))
  }

  function handlePointerDown(event: ReactPointerEvent<HTMLDivElement>) {
    event.preventDefault()
    setRatioFromClientY(event.clientY)

    const handlePointerMove = (moveEvent: PointerEvent) => {
      setRatioFromClientY(moveEvent.clientY)
    }
    const handlePointerUp = () => {
      document.removeEventListener('pointermove', handlePointerMove)
      document.removeEventListener('pointerup', handlePointerUp)
      document.body.classList.remove('resizingSplit')
    }

    document.body.classList.add('resizingSplit')
    document.addEventListener('pointermove', handlePointerMove)
    document.addEventListener('pointerup', handlePointerUp)
  }

  function handleKeyDown(event: ReactKeyboardEvent<HTMLDivElement>) {
    if (event.key === 'ArrowUp') {
      event.preventDefault()
      setRatio((current) => clampRatio(current - 0.05, minRatio, maxRatio))
    }
    if (event.key === 'ArrowDown') {
      event.preventDefault()
      setRatio((current) => clampRatio(current + 0.05, minRatio, maxRatio))
    }
    if (event.key === 'Home') {
      event.preventDefault()
      setRatio(minRatio)
    }
    if (event.key === 'End') {
      event.preventDefault()
      setRatio(maxRatio)
    }
  }

  return (
    <div
      ref={containerRef}
      className={`resizableSplit ${className ?? ''}`}
      style={{
        gridTemplateRows: `minmax(0, ${firstPercent}%) 10px minmax(0, ${secondPercent}%)`,
      }}
    >
      <div className={`splitPane ${firstClassName ?? ''}`}>{first}</div>
      <div
        className="splitHandle"
        role="separator"
        tabIndex={0}
        aria-label={ariaLabel}
        aria-orientation="horizontal"
        aria-valuemin={Math.round(minRatio * 100)}
        aria-valuemax={Math.round(maxRatio * 100)}
        aria-valuenow={Math.round(ratio * 100)}
        onPointerDown={handlePointerDown}
        onKeyDown={handleKeyDown}
      >
        <span className="splitHandleGrip" />
      </div>
      <div className={`splitPane ${secondClassName ?? ''}`}>{second}</div>
    </div>
  )
}

export function HeaderRowControl({
  headerRow,
  disabled,
  onHeaderRowChange,
  onApply,
}: {
  headerRow: number
  disabled: boolean
  onHeaderRowChange: (row: number) => void
  onApply: () => void
}) {
  return (
    <div className="headerRowControl" data-testid="header-row-control">
      <label htmlFor="header-row-input">表头行</label>
      <Input
        id="header-row-input"
        aria-label="表头行"
        className="headerRowInput"
        type="number"
        min={1}
        size="small"
        value={String(headerRow)}
        onChange={(_, data) => onHeaderRowChange(Number(data.value) || 1)}
      />
      <Button size="small" disabled={disabled} onClick={onApply}>
        应用
      </Button>
    </div>
  )
}

export function ResultActions({
  canUseResult,
  copied,
  exporting,
  onCopy,
  onExport,
}: {
  canUseResult: boolean
  copied: boolean
  exporting: boolean
  onCopy: () => void
  onExport: () => void
}) {
  return (
    <div className="resultActions">
      {copied && (
        <span className="copied">
          <CheckmarkCircle20Filled />
          已复制
        </span>
      )}
      <Button icon={<Copy24Regular />} disabled={!canUseResult} onClick={onCopy}>
        复制结果
      </Button>
      <Button
        className={`exportButton ${canUseResult ? 'ready' : ''}`}
        appearance={canUseResult ? 'primary' : 'secondary'}
        icon={<ArrowDownload24Regular />}
        disabled={!canUseResult || exporting}
        onClick={onExport}
      >
        导出
      </Button>
    </div>
  )
}

export function SheetTabs({
  sheets,
  selectedSheet,
  disabled,
  emptyText = '选择 Excel 后显示 sheet',
  onSelect,
}: {
  sheets: string[]
  selectedSheet: string
  disabled: boolean
  emptyText?: string
  onSelect: (sheet: string) => void
}) {
  return (
    <div className="sheetTabBar" aria-label="工作表标签">
      {sheets.length === 0 ? (
        <span className="sheetTabHint">{emptyText}</span>
      ) : (
        sheets.map((sheet) => (
          <button
            type="button"
            className={`sheetTab ${sheet === selectedSheet ? 'active' : ''}`}
            key={sheet}
            disabled={disabled}
            onClick={() => onSelect(sheet)}
          >
            {sheet}
          </button>
        ))
      )}
      <span className="sheetTabAdd">+</span>
    </div>
  )
}

export function GuideDialog({
  open,
  dontShowAgain,
  onDontShowAgainChange,
  onClose,
}: {
  open: boolean
  dontShowAgain: boolean
  onDontShowAgainChange: (checked: boolean) => void
  onClose: () => void
}) {
  return (
    <Dialog open={open} onOpenChange={(_, data) => { if (!data.open) onClose() }}>
      <DialogSurface className="guideSurface">
        <DialogBody>
          <DialogTitle>新手指引</DialogTitle>
          <DialogContent className="guideContent">
            <ol>
              <li>选择本地 Excel 文件，底部 sheet 标签可以切换工作表。</li>
              <li>按实际表头修改“表头行”，再选择分组列和求和列。</li>
              <li>双击原始数据可先修正内容，生成汇总时会使用修改后的数据。</li>
              <li>汇总结果也支持双击编辑，确认后可复制带格式内容或导出新的 Excel。</li>
            </ol>
            <Checkbox
              label="下次不再提示"
              checked={dontShowAgain}
              onChange={(_, data) => onDontShowAgainChange(Boolean(data.checked))}
            />
          </DialogContent>
          <DialogActions>
            <Button onClick={onClose}>跳过</Button>
            <Button appearance="primary" onClick={onClose}>开始使用</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  )
}

function ColumnChecks({
  headers,
  selected,
  disabled,
  onToggle,
}: {
  headers: string[]
  selected: string[]
  disabled: string[]
  onToggle: (column: string, checked: boolean) => void
}) {
  if (headers.length === 0) return <div className="emptyState">暂无列</div>

  return (
    <div className="checkList">
      {headers.map((header) => (
        <Checkbox
          key={header}
          label={header}
          checked={selected.includes(header)}
          disabled={disabled.includes(header)}
          onChange={(_, data) => onToggle(header, Boolean(data.checked))}
        />
      ))}
    </div>
  )
}

function ExcelWorkbookFrame({
  columns,
  rows,
  rowIndexes,
  firstRowNumber,
  headerStyle,
  dataStyle,
  rowHeights,
  emptyText,
  footer,
  headerFilters,
  filterOptionRows,
  draggableColumns = false,
  onHeaderFilterChange,
  onCellChange,
  onColumnReorder,
}: {
  columns: PreviewColumn[]
  rows: GridRow[]
  rowIndexes?: number[]
  firstRowNumber?: number
  headerStyle?: CellStyle
  dataStyle?: CellStyle
  rowHeights?: Record<string, number>
  emptyText: string
  footer: ReactNode
  headerFilters?: ColumnFilters
  filterOptionRows?: GridRow[]
  draggableColumns?: boolean
  onHeaderFilterChange?: (column: string, values: string[]) => void
  onCellChange: (rowIndex: number, column: string, value: string) => void
  onColumnReorder?: (fromKey: string, toKey: string) => void
}) {
  return (
    <div className="workbookFrame">
      <ExcelGrid
        columns={columns}
        rows={rows}
        rowIndexes={rowIndexes}
        firstRowNumber={firstRowNumber}
        headerStyle={headerStyle}
        dataStyle={dataStyle}
        rowHeights={rowHeights}
        emptyText={emptyText}
        headerFilters={headerFilters}
        filterOptionRows={filterOptionRows}
        draggableColumns={draggableColumns}
        onHeaderFilterChange={onHeaderFilterChange}
        onCellChange={onCellChange}
        onColumnReorder={onColumnReorder}
      />
      {footer}
    </div>
  )
}

function ExcelGrid({
  columns,
  rows,
  rowIndexes,
  firstRowNumber = 1,
  headerStyle,
  dataStyle,
  rowHeights,
  emptyText,
  headerFilters,
  filterOptionRows,
  draggableColumns = false,
  onHeaderFilterChange,
  onCellChange,
  onColumnReorder,
}: {
  columns: PreviewColumn[]
  rows: GridRow[]
  rowIndexes?: number[]
  firstRowNumber?: number
  headerStyle?: CellStyle
  dataStyle?: CellStyle
  rowHeights?: Record<string, number>
  emptyText: string
  headerFilters?: ColumnFilters
  filterOptionRows?: GridRow[]
  draggableColumns?: boolean
  onHeaderFilterChange?: (column: string, values: string[]) => void
  onCellChange: (rowIndex: number, column: string, value: string) => void
  onColumnReorder?: (fromKey: string, toKey: string) => void
}) {
  const [editing, setEditing] = useState<{ rowIndex: number; column: string; value: string } | null>(null)
  const [draggedColumn, setDraggedColumn] = useState<string | null>(null)
  const [openFilterColumn, setOpenFilterColumn] = useState<string | null>(null)
  const resolvedRowIndexes = rowIndexes ?? rows.map((_, index) => index)
  const maxVisibleRowNumber = Math.max(
    firstRowNumber,
    ...resolvedRowIndexes.map((sourceIndex) => firstRowNumber + sourceIndex + 1),
  )
  const rowNumberWidth = compactRowNumberWidth(maxVisibleRowNumber)
  const rowsForFilterOptions = filterOptionRows ?? rows
  const headerRowHeight = rowHeights?.['1'] ?? 24

  if (columns.length === 0) {
    return <div className="gridEmpty">{emptyText}</div>
  }

  function commitEdit() {
    if (!editing) return
    onCellChange(editing.rowIndex, editing.column, editing.value)
    setEditing(null)
  }

  function handleColumnDragStart(event: DragEvent<HTMLElement>, columnKey: string) {
    if (!draggableColumns) return
    event.dataTransfer.effectAllowed = 'move'
    setDraggedColumn(columnKey)
  }

  function handleColumnDrop(event: DragEvent<HTMLElement>, targetKey: string) {
    event.preventDefault()
    if (!draggedColumn) return
    onColumnReorder?.(draggedColumn, targetKey)
    setDraggedColumn(null)
  }

  function handleFilterToggle(column: string, option: string, checked: boolean) {
    const selected = headerFilters?.[column] ?? []
    const next = checked
      ? [...new Set([...selected, option])]
      : selected.filter((value) => value !== option)
    onHeaderFilterChange?.(column, next)
  }

  return (
    <div className="excelFrame">
      <table className="excelGrid">
        <colgroup>
          <col className="rowNumberColumn" style={{ width: `${rowNumberWidth}px` }} />
          {columns.map((column) => (
            <col key={column.key} style={{ width: `${column.width_px}px` }} />
          ))}
        </colgroup>
        <thead>
          <tr className="columnLetters">
            <th className="cornerCell" style={{ width: `${rowNumberWidth}px` }} />
            {columns.map((column, index) => (
              <th
                key={column.key}
                className={draggableColumns ? 'draggableColumnHeader' : undefined}
                draggable={draggableColumns}
                onDragStart={(event) => handleColumnDragStart(event, column.key)}
                onDragOver={(event) => {
                  if (draggedColumn) event.preventDefault()
                }}
                onDrop={(event) => handleColumnDrop(event, column.key)}
                onDragEnd={() => setDraggedColumn(null)}
              >
                {column.letter || excelColumnName(index + 1)}
              </th>
            ))}
          </tr>
          <tr style={{ height: `${headerRowHeight}px` }}>
            <th className="rowHeader" style={{ width: `${rowNumberWidth}px` }}>{firstRowNumber}</th>
            {columns.map((column) => (
              <th
                key={column.key}
                className={draggableColumns ? 'draggableColumnHeader' : undefined}
                draggable={draggableColumns}
                style={styleToReact(headerStyle, '#ffffff')}
                onDragStart={(event) => handleColumnDragStart(event, column.key)}
                onDragOver={(event) => {
                  if (draggedColumn) event.preventDefault()
                }}
                onDrop={(event) => handleColumnDrop(event, column.key)}
                onDragEnd={() => setDraggedColumn(null)}
              >
                <div className="columnHeaderContent">
                  <span className="columnTitle">{column.key}</span>
                  {onHeaderFilterChange && (
                    <HeaderFilterMenu
                      column={column.key}
                      options={filterOptionsForColumn(rowsForFilterOptions, column.key)}
                      selected={headerFilters?.[column.key] ?? []}
                      open={openFilterColumn === column.key}
                      onOpenChange={(open) => setOpenFilterColumn(open ? column.key : null)}
                      onToggle={handleFilterToggle}
                      onClear={() => onHeaderFilterChange(column.key, [])}
                    />
                  )}
                </div>
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, index) => (
            <tr key={index} style={{ height: `${rowHeights?.['2'] ?? 22}px` }}>
              <th className="rowHeader" style={{ width: `${rowNumberWidth}px` }}>
                {firstRowNumber + (resolvedRowIndexes[index] ?? index) + 1}
              </th>
              {columns.map((column) => {
                const sourceRowIndex = resolvedRowIndexes[index] ?? index
                const isEditing = editing?.rowIndex === sourceRowIndex && editing.column === column.key
                const value = row[column.key] ?? ''
                return (
                  <td
                    key={column.key}
                    className={isEditing ? 'editingCell' : ''}
                    style={styleToReact({ ...dataStyle, ...column.style }, '#ffffff')}
                    onDoubleClick={() =>
                      setEditing({
                        rowIndex: resolvedRowIndexes[index] ?? index,
                        column: column.key,
                        value: String(value),
                      })
                    }
                  >
                    {isEditing ? (
                      <input
                        className="cellEditor"
                        value={editing.value}
                        autoFocus
                        onChange={(event) => setEditing({ ...editing, value: event.target.value })}
                        onBlur={commitEdit}
                        onKeyDown={(event) => {
                          if (event.key === 'Enter') commitEdit()
                          if (event.key === 'Escape') setEditing(null)
                        }}
                      />
                    ) : (
                      value
                    )}
                  </td>
                )
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

export function HeaderFilterMenu({
  column,
  options,
  selected,
  selectedValues,
  open,
  onOpenChange,
  onToggle,
  onClear,
  onChange,
}: {
  column: string
  options: string[]
  selected?: string[]
  selectedValues?: string[]
  open?: boolean
  onOpenChange?: (open: boolean) => void
  onToggle?: (column: string, option: string, checked: boolean) => void
  onClear?: () => void
  onChange?: (values: string[]) => void
}) {
  const [localOpen, setLocalOpen] = useState(false)
  const [menuStyle, setMenuStyle] = useState<CSSProperties>({})
  const rootRef = useRef<HTMLSpanElement>(null)
  const buttonRef = useRef<HTMLButtonElement>(null)
  const isOpen = open ?? localOpen
  const selectedItems = selected ?? selectedValues ?? []

  const updateMenuPosition = useCallback(() => {
    const buttonRect = buttonRef.current?.getBoundingClientRect()
    if (!buttonRect) return

    const width = Math.min(FILTER_MENU_WIDTH, window.innerWidth - FILTER_MENU_MARGIN * 2)
    const maxHeight = Math.min(FILTER_MENU_MAX_HEIGHT, window.innerHeight * 0.45)
    const centeredLeft = buttonRect.left + buttonRect.width / 2 - width / 2
    const left = clampNumber(centeredLeft, FILTER_MENU_MARGIN, window.innerWidth - width - FILTER_MENU_MARGIN)
    const belowTop = buttonRect.bottom + FILTER_MENU_GAP
    const opensAbove = belowTop + maxHeight > window.innerHeight - FILTER_MENU_MARGIN && buttonRect.top > maxHeight
    const top = opensAbove
      ? Math.max(FILTER_MENU_MARGIN, buttonRect.top - maxHeight - FILTER_MENU_GAP)
      : Math.min(belowTop, window.innerHeight - FILTER_MENU_MARGIN)

    setMenuStyle({
      left: `${left}px`,
      top: `${top}px`,
      width: `${width}px`,
    })
  }, [])

  function setOpen(nextOpen: boolean) {
    onOpenChange?.(nextOpen)
    if (open === undefined) setLocalOpen(nextOpen)
  }

  useEffect(() => {
    if (!isOpen) return undefined

    const handleOutsidePointerDown = (event: PointerEvent) => {
      const target = event.target
      if (target instanceof Node && rootRef.current?.contains(target)) return
      onOpenChange?.(false)
      if (open === undefined) setLocalOpen(false)
    }

    document.addEventListener('pointerdown', handleOutsidePointerDown)
    return () => document.removeEventListener('pointerdown', handleOutsidePointerDown)
  }, [isOpen, onOpenChange, open])

  useLayoutEffect(() => {
    if (!isOpen) return undefined

    updateMenuPosition()
    window.addEventListener('resize', updateMenuPosition)
    window.addEventListener('scroll', updateMenuPosition, true)
    return () => {
      window.removeEventListener('resize', updateMenuPosition)
      window.removeEventListener('scroll', updateMenuPosition, true)
    }
  }, [isOpen, updateMenuPosition])

  function toggleOption(option: string, checked: boolean) {
    const next = checked
      ? [...new Set([...selectedItems, option])]
      : selectedItems.filter((value) => value !== option)
    onChange?.(next)
    onToggle?.(column, option, checked)
  }

  function clearFilter() {
    onChange?.([])
    onClear?.()
  }

  return (
    <span ref={rootRef} className="headerFilterWrap" onClick={(event) => event.stopPropagation()}>
      <button
        ref={buttonRef}
        type="button"
        className={`filterMenuButton headerFilterButton ${selectedItems.length > 0 ? 'active' : ''}`}
        aria-label={`筛选 ${column}`}
        onClick={() => setOpen(!isOpen)}
      >
        <Filter16Regular />
        {selectedItems.length > 0 && <span className="filterCount">{selectedItems.length}</span>}
      </button>
      {isOpen && (
        <div
          className="headerFilterMenu"
          style={menuStyle}
          onDoubleClick={(event) => event.stopPropagation()}
          onWheel={(event) => event.stopPropagation()}
        >
          <div className="filterMenuTop">
            <span>{column}</span>
            <button type="button" onClick={clearFilter}>清除</button>
          </div>
          <div className="filterOptionList">
            {options.map((option) => (
              <label key={option} className="filterOption">
                <input
                  type="checkbox"
                  checked={selectedItems.includes(option)}
                  onChange={(event) => toggleOption(option, event.target.checked)}
                />
                <span>{option}</span>
              </label>
            ))}
          </div>
        </div>
      )}
    </span>
  )
}

function clampNumber(value: number, min: number, max: number) {
  return Math.min(max, Math.max(min, value))
}

function clampRatio(value: number, minRatio: number, maxRatio: number) {
  return Math.min(maxRatio, Math.max(minRatio, value))
}

function formatRatioPercent(value: number) {
  const percent = Number((value * 100).toFixed(1))
  return Number.isInteger(percent) ? `${percent}` : `${percent}`
}

function styleToReact(style?: CellStyle, fallbackBackground?: string): CSSProperties {
  if (!style) return { backgroundColor: fallbackBackground }
  return {
    backgroundColor: style.background ?? fallbackBackground,
    color: style.color ?? undefined,
    fontWeight: style.bold ? 700 : undefined,
    textAlign: style.horizontal as CSSProperties['textAlign'],
    verticalAlign: style.vertical as CSSProperties['verticalAlign'],
    fontFamily: style.font_name ?? undefined,
    fontSize: style.font_size ? `${style.font_size}pt` : undefined,
    border: style.border_style ? `1px solid ${style.border_color ?? '#d1d5db'}` : undefined,
  }
}

function guessColumns(headers: string[]) {
  const group = headers.filter((header) =>
    ['规格型号', '型号', '单位', '单价'].some((keyword) => header.includes(keyword)),
  )
  const sum = headers.filter((header) =>
    ['数量', '金额'].some((keyword) => header.includes(keyword)),
  )
  return {
    group,
    sum: sum.filter((header) => !group.includes(header)),
  }
}

function excelColumnName(index: number) {
  let name = ''
  let current = index
  while (current > 0) {
    const remainder = (current - 1) % 26
    name = String.fromCharCode(65 + remainder) + name
    current = Math.floor((current - 1) / 26)
  }
  return name
}

async function saveBlobWithPicker(blob: Blob, fileName: string) {
  const picker = (window as WindowWithFilePicker).showSaveFilePicker
  if (picker) {
    const handle = await picker({
      suggestedName: fileName,
      types: [
        {
          description: 'Excel 工作簿',
          accept: {
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
          },
        },
      ],
    })
    const writable = await handle.createWritable()
    await writable.write(blob)
    await writable.close()
    return
  }

  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = fileName
  link.click()
  URL.revokeObjectURL(url)
}

type WindowWithFilePicker = Window & {
  showSaveFilePicker?: (options: {
    suggestedName?: string
    types?: Array<{
      description: string
      accept: Record<string, string[]>
    }>
  }) => Promise<{
    createWritable: () => Promise<{
      write: (data: Blob) => Promise<void>
      close: () => Promise<void>
    }>
  }>
}

function loadingLabel(key: string) {
  const labels: Record<string, string> = {
    upload: '正在读取 Excel',
    preview: '正在刷新预览',
    summary: '正在生成汇总',
    export: '正在导出 Excel',
  }
  return labels[key] ?? '处理中'
}

function errorMessage(error: unknown) {
  return error instanceof Error ? error.message : '操作失败'
}

function shouldShowGuide() {
  try {
    return window.localStorage.getItem(GUIDE_STORAGE_KEY) !== '1'
  } catch {
    return true
  }
}

function rememberGuidePreference(dontShowAgain: boolean) {
  if (!dontShowAgain) return
  try {
    window.localStorage.setItem(GUIDE_STORAGE_KEY, '1')
  } catch {
    // Ignore storage failures; the guide can still be skipped for this session.
  }
}

export default App
