/**
 * Jest setup file: polyfills and global stubs required by Fluent UI v9 components
 * in a jsdom (Node.js) test environment.
 */

// Fluent UI v9's MessageBar uses ResizeObserver internally via useMessageBarReflow.
// jsdom does not implement ResizeObserver, so we provide a no-op stub.
if (typeof globalThis.ResizeObserver === 'undefined') {
  globalThis.ResizeObserver = class ResizeObserver {
    observe() { /* no-op in test environment */ }
    unobserve() { /* no-op in test environment */ }
    disconnect() { /* no-op in test environment */ }
  };
}
