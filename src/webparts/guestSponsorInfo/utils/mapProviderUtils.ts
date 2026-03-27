// SPDX-FileCopyrightText: 2026 Workoho GmbH <https://workoho.com>
// SPDX-FileCopyrightText: 2026 Julian Pawlowski <https://github.com/jpawlowski>
// SPDX-License-Identifier: LicenseRef-PolyForm-Shield-1.0.0

/**
 * Map provider type.
 */
export type MapProvider = 'bing' | 'google' | 'apple' | 'openstreetmap' | 'none';

/**
 * Map provider configuration for auto mode.
 */
export interface MapProviderConfig {
  mode: 'auto' | 'manual' | 'none';
  /** Used in manual mode or as fallback */
  manualProvider?: MapProvider;
  /** OS-specific providers (auto mode) */
  iosProvider?: MapProvider;
  androidProvider?: MapProvider;
  windowsProvider?: MapProvider;
  macosProvider?: MapProvider;
  linuxProvider?: MapProvider;
}

/**
 * Detected operating system.
 */
export type DetectedOS = 'iOS' | 'Android' | 'Windows' | 'macOS' | 'Linux' | 'unknown';

/**
 * Detect the user's operating system from the User-Agent string.
 * Prioritizes mobile OS detection before desktop.
 */
export function detectOS(userAgent: string): DetectedOS {
  // Mobile: Check first (more specific patterns)
  if (/iPhone|iPad|iPod/.test(userAgent)) {
    return 'iOS';
  }
  if (/Android/.test(userAgent)) {
    return 'Android';
  }

  // Desktop
  if (/Windows/.test(userAgent)) {
    return 'Windows';
  }
  if (/Macintosh.*Intel|Macintosh.*PPC|Macintosh.*AppleSilicon/.test(userAgent)) {
    return 'macOS';
  }
  if (/Linux/.test(userAgent) && !/Android/.test(userAgent)) {
    return 'Linux';
  }

  return 'unknown';
}

/**
 * Determine the effective map provider based on the user's OS and the configured settings.
 * In manual mode, returns the manual provider (applies to all OS).
 * In auto mode, returns the provider configured for the detected OS.
 * Falls back to sensible per-OS defaults when no explicit provider is configured.
 */
export function getEffectiveMapProvider(
  userAgent: string,
  config: MapProviderConfig
): MapProvider {
  // None mode: map links disabled globally
  if (config.mode === 'none') {
    return 'none';
  }

  // Manual mode: same provider for all users
  if (config.mode === 'manual') {
    return config.manualProvider ?? 'bing';
  }

  // Auto mode: OS-specific selection
  const os = detectOS(userAgent);

  switch (os) {
    case 'iOS':
      return config.iosProvider ?? 'apple';
    case 'Android':
      return config.androidProvider ?? 'google';
    case 'Windows':
      return config.windowsProvider ?? 'bing';
    case 'macOS':
      return config.macosProvider ?? 'apple';
    case 'Linux':
      return config.linuxProvider ?? 'openstreetmap';
    case 'unknown':
    default:
      // Fallback for unknown OS: use Linux default
      return config.linuxProvider ?? 'openstreetmap';
  }
}

/**
 * Build an external map link for the given provider and address query.
 */
export function buildExternalMapLink(
  provider: MapProvider,
  address: string
): string | undefined {
  if (provider === 'none') {
    return undefined;
  }

  const query = encodeURIComponent(address);

  switch (provider) {
    case 'google':
      return `https://www.google.com/maps/search/?api=1&query=${query}`;
    case 'apple':
      return `https://maps.apple.com/?q=${query}`;
    case 'openstreetmap':
      return `https://www.openstreetmap.org/search?query=${query}`;
    case 'bing':
    default:
      return `https://www.bing.com/maps?q=${query}`;
  }
}
