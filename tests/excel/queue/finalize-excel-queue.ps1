[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDirectory,
    [Parameter(Mandatory = $true)]
    [string]$Result,
    [string]$CompletionMarker = 'QUEUE_COMPLETE'
)

$ErrorActionPreference = 'Stop'
$sourceDirectoryPath = (Resolve-Path -LiteralPath $SourceDirectory -ErrorAction Stop).Path
$resultPath = (Resolve-Path -LiteralPath $Result -ErrorAction Stop).Path
$resultName = [System.IO.Path]::GetFileName($resultPath)
$sourcePrefix = $null
if ($resultName.StartsWith('ExcelQueue', [System.StringComparison]::OrdinalIgnoreCase)) {
    $sourcePrefix = 'ExcelQueue'
} elseif ($resultName.StartsWith('FormatMatrix', [System.StringComparison]::OrdinalIgnoreCase)) {
    $sourcePrefix = 'FormatMatrix'
} elseif ($resultName.StartsWith('RadixMatrix', [System.StringComparison]::OrdinalIgnoreCase)) {
    $sourcePrefix = 'RadixMatrix'
}
if ($null -eq $sourcePrefix) {
    throw "Cannot determine VBA source group from result filename: $resultName"
}
$utf8 = New-Object System.Text.UTF8Encoding($false)
$resultText = [System.IO.File]::ReadAllText($resultPath, $utf8)

if ($resultText -notmatch "(?m)^$([regex]::Escape($CompletionMarker))=True\s*$") {
    throw "Excel queue did not reach its completion marker ($CompletionMarker): $resultPath"
}

$sourceFiles = Get-ChildItem -LiteralPath $sourceDirectoryPath -File |
    Where-Object {
        $_.Extension -match '^\.(bas|cls|frm)$' -and
        $_.Name.StartsWith($sourcePrefix, [System.StringComparison]::OrdinalIgnoreCase)
    } |
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
$hash = ([System.BitConverter]::ToString($hashBytes)).Replace('-', '').ToLowerInvariant()
$recordedHashMatch = [regex]::Match($resultText, '(?mi)^QUEUE_SOURCE_SHA256=([0-9a-f]{64})\s*$')
if (-not $recordedHashMatch.Success) {
    throw "Excel queue result is missing the embedded workbook source hash: $resultPath"
}
if ($recordedHashMatch.Groups[1].Value.ToLowerInvariant() -ne $hash) {
    throw "Excel queue result source hash does not match the current source: $resultPath (recorded=$($recordedHashMatch.Groups[1].Value.ToLowerInvariant()), current=$hash, group=$sourcePrefix)"
}
Write-Output "Verified Excel queue result: $resultPath"
