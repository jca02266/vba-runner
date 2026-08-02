[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path
)

$ErrorActionPreference = 'Stop'
$resolved = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
    throw "File not found: $Path"
}

# VBA Print writes using the Windows system code page. Rewrite the completed
# file as UTF-8 without a BOM for stable repository comparison.
$text = [System.IO.File]::ReadAllText($resolved, [System.Text.Encoding]::Default)
[System.IO.File]::WriteAllText($resolved, $text, (New-Object System.Text.UTF8Encoding($false)))
Write-Output "Converted to UTF-8 without BOM: $resolved"
