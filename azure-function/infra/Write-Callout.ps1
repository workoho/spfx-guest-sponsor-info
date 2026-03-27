<#
.SYNOPSIS
    Callout box helpers for PowerShell scripts — the PowerShell equivalent
    of the hint / next_steps / important functions in scripts/colors.sh.

.DESCRIPTION
    Dot-source this file at the top of any PowerShell script that needs
    developer-facing callout boxes:

        . "$PSScriptRoot/Write-Callout.ps1"

    Then call Write-Hint, Write-NextSteps, or Write-Important with one or
    more lines of text. Pass $null or '' for a blank separator line.

.EXAMPLE
    Write-Hint 'Edit .env and set SPFX_SERVE_TENANT_DOMAIN'

.EXAMPLE
    Write-Important `
        'Edit azure-function/local.settings.json' `
        '' `
        'Required:' `
        '  TENANT_ID — your Entra tenant ID'

.EXAMPLE
    Write-NextSteps `
        'Paste these values into the SPFx web part property pane:' `
        '' `
        "  Sponsor API URL   : $url" `
        "  Function Client ID: $clientId"

.NOTES
    Copyright 2026 Workoho GmbH <https://workoho.com>
    Author: Julian Pawlowski <https://github.com/jpawlowski>
    Licensed under PolyForm Shield License 1.0.0
    <https://polyformproject.org/licenses/shield/1.0.0>
#>

function Write-Box {
    <#
    .SYNOPSIS
        Internal: draws a coloured callout box around lines of text.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][ConsoleColor]$Color,
        [Parameter(ValueFromRemainingArguments)][string[]]$Lines
    )

    $rule = '─' * 59
    $tlen = $Title.Length
    $dashes = 56 - $tlen
    if ($dashes -lt 4) { $dashes = 4 }
    $headerDashes = '─' * $dashes
    $footerDashes = '─' * 59

    Write-Host ''
    Write-Host "  ╭─ " -ForegroundColor $Color -NoNewline
    Write-Host $Title -ForegroundColor $Color -NoNewline
    Write-Host " $headerDashes" -ForegroundColor $Color
    Write-Host "  │" -ForegroundColor $Color

    foreach ($line in $Lines) {
        if ([string]::IsNullOrEmpty($line)) {
            Write-Host "  │" -ForegroundColor $Color
        }
        else {
            Write-Host "  │" -ForegroundColor $Color -NoNewline
            Write-Host "  $line"
        }
    }

    Write-Host "  │" -ForegroundColor $Color
    Write-Host "  ╰$footerDashes" -ForegroundColor $Color
    Write-Host ''
}

function Write-Hint {
    <# .SYNOPSIS Cyan callout — developer tips, good-to-know info. #>
    [CmdletBinding()]
    param([Parameter(ValueFromRemainingArguments)][string[]]$Lines)
    Write-Box -Title 'HINT' -Color Cyan @Lines
}

function Write-NextSteps {
    <# .SYNOPSIS Green callout — what to do after the script finishes. #>
    [CmdletBinding()]
    param([Parameter(ValueFromRemainingArguments)][string[]]$Lines)
    Write-Box -Title 'NEXT STEPS' -Color Green @Lines
}

function Write-Important {
    <# .SYNOPSIS Yellow callout — critical action items that must be done. #>
    [CmdletBinding()]
    param([Parameter(ValueFromRemainingArguments)][string[]]$Lines)
    Write-Box -Title 'IMPORTANT' -Color Yellow @Lines
}
