#!/usr/bin/env -S pwsh -NoLogo -NoProfile
# Copyright 2026 Workoho GmbH <https://workoho.com>
# Author: Julian Pawlowski <https://github.com/jpawlowski>
# Licensed under PolyForm Shield License 1.0.0 <https://polyformproject.org/licenses/shield/1.0.0>
#
# Run PSScriptAnalyzer for all PowerShell scripts under azure-function/infra/.
#
# Usage:
#   pwsh -NonInteractive -File scripts/lint-ps.ps1
#   npm run lint:ps

$ErrorActionPreference = 'Stop'

# Always run from the repository root so the settings file path resolves correctly.
Set-Location (Join-Path $PSScriptRoot '..')

$results = Invoke-ScriptAnalyzer -Path 'azure-function/infra' -Recurse `
    -Settings 'PSScriptAnalyzerSettings.psd1'

if ($results) {
    $results | Format-Table RuleName, Severity, Line, ScriptName, Message -AutoSize
    exit 1
}
