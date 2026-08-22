[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDirectory,
    [Parameter(Mandatory = $true)]
    [string]$Result,
    [Parameter(Mandatory = $true)]
    [string]$CompletionMarker,
    [Parameter(Mandatory = $true)]
    [string]$SourcePrefix
)

$ErrorActionPreference = 'Stop'

function Test-CurrentExcelResult {
    $sourceDirectoryPath = (Resolve-Path -LiteralPath $SourceDirectory -ErrorAction Stop).Path
    if (-not (Test-Path -LiteralPath $Result -PathType Leaf)) {
        return $false
    }
    $resultPath = (Resolve-Path -LiteralPath $Result -ErrorAction Stop).Path
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $resultText = [System.IO.File]::ReadAllText($resultPath, $utf8)
    $markerPattern = '(?m)^' + [regex]::Escape($CompletionMarker) + '=True\s*$'
    if ($resultText -notmatch $markerPattern) {
        return $false
    }
    $recorded = [regex]::Match($resultText, '(?mi)^QUEUE_SOURCE_SHA256=([0-9a-f]{64})\s*$')
    if (-not $recorded.Success) {
        return $false
    }
    $sourceFiles = @(Get-ChildItem -LiteralPath $sourceDirectoryPath -File |
        Where-Object {
            $_.Extension -match '^\.(bas|cls|frm)$' -and
            $_.Name.StartsWith($SourcePrefix, [System.StringComparison]::OrdinalIgnoreCase)
        } |
        Sort-Object -Property Name)
    if ($sourceFiles.Count -eq 0) {
        throw "No VBA source files found for group $($SourcePrefix): $sourceDirectoryPath"
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
    return $recorded.Groups[1].Value.ToLowerInvariant() -eq $current
}

try {
    $isCurrent = Test-CurrentExcelResult
} catch {
    Write-Error $_
    exit 2
}
if ($isCurrent) {
    Write-Output "Excel result is current for $($SourcePrefix): $Result"
    exit 0
}
exit 1
