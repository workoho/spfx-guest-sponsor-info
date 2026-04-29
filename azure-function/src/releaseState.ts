// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0

/**
 * In-memory release state shared between the checkGitHubRelease timer trigger
 * and the getGuestSponsors request handler.
 *
 * Node.js module state persists for the lifetime of a function instance, i.e.
 * across all invocations within the same cold start.  On a fresh cold start
 * `latestGitHubVersion` is `undefined` until `checkGitHubRelease` has completed
 * at least one successful GitHub API call (`runOnStartup: true` guarantees this
 * happens shortly after every deployment or restart).
 *
 * Thread safety: Node.js is single-threaded; concurrent async invocations share
 * the same call stack, so plain module-level `let` variables are safe here.
 */

/**
 * The latest published GitHub release version (semver without leading "v"),
 * updated by `checkGitHubRelease` after each successful GitHub API call —
 * regardless of whether the version is newer than the current function version.
 * `undefined` until the timer has completed its first run in this instance.
 */
export let latestGitHubVersion: string | undefined;

/**
 * The HTML URL of the latest published GitHub release page
 * (for example, a GitHub Releases tag page).
 * Updated alongside `latestGitHubVersion` by `checkGitHubRelease`.
 * `undefined` until the timer has completed its first run in this instance.
 */
export let latestGitHubReleaseUrl: string | undefined;

/**
 * Records the most recent GitHub release version fetched by the timer.
 * Called by `checkGitHubRelease` after every successful GitHub API response.
 */
export function setLatestGitHubVersion(version: string): void {
  latestGitHubVersion = version;
}

/**
 * Records the HTML URL of the most recent GitHub release page.
 * Called by `checkGitHubRelease` alongside `setLatestGitHubVersion`.
 */
export function setLatestGitHubReleaseUrl(url: string): void {
  latestGitHubReleaseUrl = url;
}

/**
 * Returns true when `candidate` is strictly newer than `current`.
 * Compares major · minor · patch only; pre-release suffixes are ignored.
 */
export function isNewerVersion(candidate: string, current: string): boolean {
  const parse = (v: string): number[] => v.split('.').map(n => parseInt(n, 10) || 0);
  const a = parse(candidate);
  const b = parse(current);
  for (let i = 0; i < 3; i++) {
    if ((a[i] ?? 0) > (b[i] ?? 0)) return true;
    if ((a[i] ?? 0) < (b[i] ?? 0)) return false;
  }
  return false;
}
