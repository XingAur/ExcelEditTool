// @vitest-environment jsdom

import { fireEvent, render, screen } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'

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
  SheetTabs,
} from './App'

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
})
