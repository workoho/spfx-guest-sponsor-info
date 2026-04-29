// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0

import { app, InvocationContext, Timer } from '@azure/functions';
import packageJson from '../package.json';
import { isNewerVersion, setLatestGitHubVersion, setLatestGitHubReleaseUrl } from './releaseState.js';

const CURRENT_VERSION: string = packageJson.version;
const GITHUB_API_URL = 'https://api.github.com/repos/workoho/spfx-guest-sponsor-info/releases/latest';

/**
 * Fetches the latest published GitHub Release and compares it against the
 * deployed function version.
 *
 * When a newer release is found, logs a structured WARNING trace that is
 * picked up by the Application Insights KQL alert rule defined in
 * monitoring.bicep.  The trace format is intentionally stable:
 *
 *   [NEW_RELEASE_AVAILABLE] currentVersion=X latestVersion=Y url=Z
 *
 * Do NOT change the key=value tokens without updating the corresponding
 * KQL extract() expressions in monitoring.bicep.
 *
 * Runs on a timer every 6 hours so the trace is regularly refreshed in
 * Application Insights and the alert stays in "fired" state until the
 * function is updated.  When the function is restarted at the same or
 * newer version the trace stops appearing, the alert auto-mitigates, and
 * —if yet another GitHub release subsequently arrives— a fresh notification
 * is sent via the info action group.
 *
 * runOnStartup: true ensures the check runs immediately after every
 * deployment/restart so admins do not have to wait up to 6 hours for the
 * first log entry following a new release.
 */
async function checkGitHubRelease(_timer: Timer, context: InvocationContext): Promise<void> {
  const currentVersion = CURRENT_VERSION.split('.').slice(0, 3).join('.');

  let response: Response;
  try {
    response = await fetch(GITHUB_API_URL, {
      headers: {
        Accept: 'application/vnd.github+json',
        // Identify our caller to GitHub for rate-limit attribution / logging.
        'User-Agent': `guest-sponsor-info-function/${currentVersion}`,
      },
      // AbortSignal.timeout() is available in Node.js 17.3+; we target 22.x.
      signal: AbortSignal.timeout(10_000),
    });
  } catch (err) {
    // Network error or timeout — log at info level (transient; not actionable).
    context.log(`[checkGitHubRelease] GitHub API unreachable — skipping: ${err}`);
    return;
  }

  if (response.status === 404) {
    // No release published yet — not an error.
    context.log('[checkGitHubRelease] No GitHub release found — repository may be pre-release.');
    return;
  }

  if (!response.ok) {
    // Rate-limit (429) or unexpected server error — transient; skip silently.
    context.log(`[checkGitHubRelease] GitHub API returned ${response.status} — skipping.`);
    return;
  }

  let data: { tag_name?: unknown; html_url?: unknown };
  try {
    data = await response.json() as { tag_name?: unknown; html_url?: unknown };
  } catch {
    context.log('[checkGitHubRelease] Could not parse GitHub API response — skipping.');
    return;
  }

  const tag = data.tag_name;
  const releaseUrl = data.html_url;
  if (typeof tag !== 'string' || !tag) return;
  if (typeof releaseUrl !== 'string' || !releaseUrl) return;

  const latestVersion = tag.replace(/^v/, '');

  // Always record the fetched GitHub version and release URL in shared in-memory
  // state so that both the getGuestSponsors mismatch-context logic and the new
  // getLatestRelease endpoint can serve clients without an extra GitHub API call.
  setLatestGitHubVersion(latestVersion);
  setLatestGitHubReleaseUrl(releaseUrl);

  if (!isNewerVersion(latestVersion, currentVersion)) {
    context.log(`[checkGitHubRelease] Function is up to date (v${currentVersion}).`);
    return;
  }

  // Structured warning — consumed by the KQL alert rule in monitoring.bicep.
  // The alert fires once per unique `latestVersion` value (dimension split)
  // and auto-mitigates once the function is updated or a newer release appears.
  context.warn(
    `[NEW_RELEASE_AVAILABLE] currentVersion=${currentVersion} latestVersion=${latestVersion} url=${releaseUrl}`
  );
}

// NCRONTAB seconds format: "ss mm hh dd MM dow"
// "0 0 */6 * * *" = every 6 hours on the hour (00:00, 06:00, 12:00, 18:00 UTC).
app.timer('checkGitHubRelease', {
  schedule: '0 0 */6 * * *',
  runOnStartup: true,
  handler: checkGitHubRelease,
});
