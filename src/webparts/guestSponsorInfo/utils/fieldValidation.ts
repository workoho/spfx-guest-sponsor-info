// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0

/**
 * Shared field-format validators used by both the WelcomeDialog wizard and
 * the property pane settings panel.  Keeping them here prevents the same
 * regex from diverging between the two entry points.
 */

/**
 * Returns true when `raw` looks like a plausible Azure Function base URL.
 *
 * Rules (intentionally permissive per product spec):
 * - `https://` / `http://` prefix is optional and stripped before checking.
 * - A trailing `/` is silently ignored.
 * - The remaining string must contain at least one dot (hostname requires a TLD
 *   or at least a second label) and must not contain whitespace.
 * - An optional port (`:8080`) and path suffix are allowed.
 *
 * Note: the stored property value already has the scheme stripped
 * (`GuestSponsorInfoWebPart.ts` strips it on save), so passing either the raw
 * user input or the stored value works correctly.
 */
export function isValidFunctionUrl(raw: string): boolean {
  // Strip protocol and trailing slash, then check what is left.
  const stripped = raw.replace(/^https?:\/\//i, '').replace(/\/+$/, '').trim();
  if (!stripped) return false;
  // hostname[.tld][/path] — at least one dot required for TLD; no bare spaces.
  return /^[A-Za-z0-9][A-Za-z0-9\-.]*\.[A-Za-z]{2,}(:\d{1,5})?(\/\S*)?$/.test(stripped);
}

/**
 * Returns true when `raw` matches the canonical GUID format
 * `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` (case-insensitive).
 */
export function isValidGuid(raw: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(raw.trim());
}
