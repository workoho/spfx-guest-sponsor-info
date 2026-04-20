@{
  ExcludeRules = @(
    # Admin and interactive setup scripts use Write-Host intentionally for
    # coloured terminal output. Write-Output does not support colours;
    # using it would require redirecting to the console host explicitly.
    'PSAvoidUsingWriteHost',

    # Scripts are UTF-8 without BOM by design: the BOM (U+FEFF) breaks
    # [scriptblock]::Create() when scripts are fetched via iwr/irm from
    # raw.githubusercontent.com. PowerShell 7 parses UTF-8 without BOM
    # correctly. Non-ASCII characters in comments are accepted.
    'PSUseBOMForUnicodeEncodedFile'
  )
}
