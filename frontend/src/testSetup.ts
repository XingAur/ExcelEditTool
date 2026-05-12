import '@testing-library/jest-dom/vitest'

const nodeFilterShim = {
  SHOW_ELEMENT: 1,
  FILTER_ACCEPT: 1,
  FILTER_REJECT: 2,
  FILTER_SKIP: 3,
}

if (!globalThis.NodeFilter) {
  Object.defineProperty(globalThis, 'NodeFilter', {
    configurable: true,
    value: nodeFilterShim,
  })
}

if (typeof window !== 'undefined' && !window.NodeFilter) {
  Object.defineProperty(window, 'NodeFilter', {
    configurable: true,
    value: nodeFilterShim,
  })
}
