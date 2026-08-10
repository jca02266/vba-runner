[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDirectory,
    [Parameter(Mandatory = $true)]
    [string]$PreparationStamp
)

$ErrorActionPreference = 'Stop'
$sourceDirectoryPath = (Resolve-Path -LiteralPath $SourceDirectory -ErrorAction Stop).Path
$stampPath = (Resolve-Path -LiteralPath $PreparationStamp -ErrorAction Stop).Path
$stamp = ([System.IO.File]::ReadAllText($stampPath)).Trim().ToLowerInvariant()
if ($stamp -notmatch '^[0-9a-f]{64}$') {
    throw "Invalid Excel queue preparation stamp: $stampPath"
}

$utf8 = New-Object System.Text.UTF8Encoding($false)
$sourceFiles = Get-ChildItem -LiteralPath $sourceDirectoryPath -File |
    Where-Object { $_.Extension -match '^\.(bas|cls|frm)$' } |
    Sort-Object -Property Name
if ($sourceFiles.Count -eq 0) {
    throw "No VBA source files found: $sourceDirectoryPath"
}
$manifest = New-Object System.Text.StringBuilder
foreach ($sourceFile in $sourceFiles) {
    $sourceText = [System.IO.File]::ReadAllText($sourceFile.FullName, $utf8).Replace("`r`n", "`n")
    $sourceText = [regex]::Replace($sourceText,
        '(?mi)^\s*Private Const QUEUE_SOURCE_SHA256\s+As String\s*=\s*"[^"]*"\s*$',
        'Private Const QUEUE_SOURCE_SHA256 As String = "__QUEUE_SOURCE_SHA256__"')
    [void]$manifest.Append($sourceFile.Name).Append("`n").Append($sourceText).Append("`n")
}
$sha256 = [System.Security.Cryptography.SHA256]::Create()
try {
    $hashBytes = $sha256.ComputeHash($utf8.GetBytes($manifest.ToString()))
} finally {
    $sha256.Dispose()
}
$current = ([System.BitConverter]::ToString($hashBytes)).Replace('-', '').ToLowerInvariant()
if ($stamp -ne $current) {
    throw "Excel queue workbook is stale. Run prepare-excel-vba.sh before copying t.xlsm to Windows. Prepared=$stamp Current=$current"
}
Write-Output "Excel queue preparation stamp is current: $current"
