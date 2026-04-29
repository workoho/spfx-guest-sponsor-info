// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0

/**
 * Jest setup file: polyfills and global stubs required by Fluent UI v9 components
 * in a jsdom (Node.js) test environment.
 */

// Suppress Griffel's safeInsertRule console.error for @container queries.
// jsdom's CSSOM (cssom library) does not support @container rules, so Griffel
// catches the CSSStyleSheet.insertRule() error and logs it. The styling is
// still applied correctly in a real browser — this is test-environment noise.
const _origConsoleError = console.error;
console.error = (...args) => {
  const msg = typeof args[0] === 'string' ? args[0] : '';
  if (msg.includes('There was a problem inserting the following rule')) return;
  _origConsoleError(...args);
};

// Fluent UI v9's MessageBar uses ResizeObserver internally via useMessageBarReflow.
// jsdom does not implement ResizeObserver, so we provide a no-op stub.
if (typeof globalThis.ResizeObserver === 'undefined') {
  globalThis.ResizeObserver = class ResizeObserver {
    observe() { /* no-op in test environment */ }
    unobserve() { /* no-op in test environment */ }
    disconnect() { /* no-op in test environment */ }
  };
}
