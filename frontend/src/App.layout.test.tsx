// @vitest-environment jsdom

import { cleanup, fireEvent, render, screen } from '@testing-library/react'
import { afterEach, describe, expect, it, vi } from 'vitest'

vi.mock('@fluentui/react-components', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@fluentui/react-components')>()
  const React = await import('react')
  const passthrough = ({ children }: { children?: React.ReactNode }) =>
    React.createElement('div', null, children)

  return {
    ...actual,
    Dialog: passthrough,
    DialogBody: passthrough,
    DialogContent: ({ children, className }: { children?: React.ReactNode; className?: string }) =>
      React.createElement('div', { className }, children),
    DialogSurface: ({ children, className }: { children?: React.ReactNode; className?: string }) =>
      React.createElement('div', { className, role: 'dialog' }, children),
    DialogTitle: ({ children }: { children?: React.ReactNode }) =>
      React.createElement('h2', null, children),
    DialogActions: passthrough,
  }
})

import { FluentProvider, webLightTheme } from '@fluentui/react-components'
import {
  GuideDialog,
  HeaderFilterMenu,
  HeaderRowControl,
  ResultActions,
  ResizableSplit,
  SheetTabs,
} from './App'

afterEach(() => cleanup())

describe('Excel-like layout refinements', () => {
  it('renders a compact inline header row control', () => {
    render(
      <FluentProvider theme={webLightTheme}>
        <HeaderRowControl
          headerRow={2}
          disabled={false}
          onHeaderRowChange={vi.fn()}
          onApply={vi.fn()}
        />
      </FluentProvider>,
    )

    expect(screen.getByText('表头行')).toBeInTheDocument()
    expect(screen.getByTestId('header-row-control')).toHaveClass('headerRowControl')
    expect(screen.getByLabelText('表头行')).toBeInTheDocument()
  })

  it('renders sheet tabs as a bottom spreadsheet tab strip', () => {
    render(
      <SheetTabs
        sheets={['上海', '安徽']}
        selectedSheet="上海"
        disabled={false}
        onSelect={vi.fn()}
      />,
    )

    expect(screen.getByLabelText('工作表标签')).toHaveClass('sheetTabBar')
    expect(screen.getByRole('button', { name: '上海' })).toHaveClass('active')
  })

  it('places export action after copy action in result toolbar', () => {
    render(
      <FluentProvider theme={webLightTheme}>
        <ResultActions
          canUseResult
          copied={false}
          exporting={false}
          onCopy={vi.fn()}
          onExport={vi.fn()}
        />
      </FluentProvider>,
    )

    const copy = screen.getByRole('button', { name: '复制结果' })
    const exportButton = screen.getByRole('button', { name: '导出' })
    expect(copy.compareDocumentPosition(exportButton)).toBe(Node.DOCUMENT_POSITION_FOLLOWING)
  })

  it('renders the first-run guide with skip and opt-out controls', () => {
    render(
      <FluentProvider theme={webLightTheme}>
        <GuideDialog
          open
          dontShowAgain={false}
          onDontShowAgainChange={vi.fn()}
          onClose={vi.fn()}
        />
      </FluentProvider>,
    )

    expect(screen.getByText('新手指引')).toBeInTheDocument()
    expect(screen.getByRole('checkbox', { name: '下次不再提示' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: '跳过' })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: '开始使用' })).toBeInTheDocument()
  })

  it('opens a compact multi-select filter menu from the column header', () => {
    const onChange = vi.fn()
    render(
      <HeaderFilterMenu
        column="型号"
        options={['A-100', 'B-200']}
        selectedValues={['A-100']}
        onChange={onChange}
      />,
    )

    const button = screen.getByRole('button', { name: '筛选 型号' })
    expect(button).toHaveClass('filterMenuButton')

    fireEvent.click(button)
    expect(screen.getByRole('checkbox', { name: 'A-100' })).toBeChecked()
    fireEvent.click(screen.getByRole('checkbox', { name: 'B-200' }))
    expect(onChange).toHaveBeenCalledWith(['A-100', 'B-200'])
  })

  it('keeps a left-edge filter menu inside the viewport', () => {
    const { container } = render(
      <HeaderFilterMenu
        column="鍨嬪彿"
        options={['A-100']}
        selectedValues={[]}
        onChange={vi.fn()}
      />,
    )
    const button = container.querySelector('.filterMenuButton') as HTMLButtonElement
    button.getBoundingClientRect = () => ({
      x: 4,
      y: 2,
      top: 2,
      left: 4,
      right: 22,
      bottom: 20,
      width: 18,
      height: 18,
      toJSON: () => ({}),
    })

    fireEvent.click(button)

    const menu = container.querySelector('.headerFilterMenu') as HTMLElement
    expect(menu.style.left).toBe('8px')
    expect(menu.style.top).toBe('26px')
    expect(menu.style.width).toBe('224px')
  })

  it('closes the filter menu when clicking outside', () => {
    render(
      <HeaderFilterMenu
        column="型号"
        options={['A-100']}
        selectedValues={[]}
        onChange={vi.fn()}
      />,
    )

    fireEvent.click(screen.getByRole('button', { name: '筛选 型号' }))
    expect(screen.getByRole('checkbox', { name: 'A-100' })).toBeInTheDocument()

    fireEvent.pointerDown(document.body)
    expect(screen.queryByRole('checkbox', { name: 'A-100' })).not.toBeInTheDocument()
  })

  it('renders a draggable split handle for resizable panes', () => {
    const { container } = render(
      <ResizableSplit
        ariaLabel="调整预览区域比例"
        first={<div>原始数据预览</div>}
        second={<div>汇总结果预览</div>}
      />,
    )

    expect(screen.getByText('原始数据预览')).toBeInTheDocument()
    expect(screen.getByText('汇总结果预览')).toBeInTheDocument()
    expect(screen.getByRole('separator', { name: '调整预览区域比例' })).toBeInTheDocument()
    expect(container.querySelector('.resizableSplit')).toHaveStyle({
      gridTemplateRows: 'minmax(0, 50%) 10px minmax(0, 50%)',
    })
  })

  it('updates the split ratio while dragging the handle', () => {
    const { container } = render(
      <ResizableSplit
        ariaLabel="调整分组列和求和列比例"
        first={<div>分组列</div>}
        second={<div>求和列</div>}
      />,
    )
    const split = container.querySelector('.resizableSplit') as HTMLElement
    split.getBoundingClientRect = () => ({
      x: 0,
      y: 0,
      top: 0,
      left: 0,
      right: 300,
      bottom: 400,
      width: 300,
      height: 400,
      toJSON: () => ({}),
    })

    const handle = screen.getByRole('separator', { name: '调整分组列和求和列比例' })
    fireEvent.pointerDown(handle, { clientY: 200 })
    fireEvent.pointerMove(document, { clientY: 300 })
    fireEvent.pointerUp(document)

    expect(split).toHaveStyle({
      gridTemplateRows: 'minmax(0, 75%) 10px minmax(0, 25%)',
    })
  })
})
