[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Source,
    [Parameter(Mandatory = $true)]
    [string]$Result
)

$ErrorActionPreference = 'Stop'
$sourcePath = (Resolve-Path -LiteralPath $Source -ErrorAction Stop).Path
$resultPath = (Resolve-Path -LiteralPath $Result -ErrorAction Stop).Path
$utf8 = New-Object System.Text.UTF8Encoding($false)
$resultText = [System.IO.File]::ReadAllText($resultPath, $utf8)

if ($resultText -notmatch '(?m)^QUEUE_COMPLETE=True\s*$') {
    throw "Excel queue did not reach its completion marker: $resultPath"
}

$sourceText = [System.IO.File]::ReadAllText($sourcePath, $utf8).Replace("`r`n", "`n")
$sha256 = [System.Security.Cryptography.SHA256]::Create()
try {
    $hashBytes = $sha256.ComputeHash($utf8.GetBytes($sourceText))
} finally {
    $sha256.Dispose()
}
$hash = ([System.BitConverter]::ToString($hashBytes)).Replace('-', '').ToLowerInvariant()
$resultText = $resultText.TrimEnd("`r", "`n") + "`nQUEUE_SOURCE_SHA256=$hash`n"
[System.IO.File]::WriteAllText($resultPath, $resultText, $utf8)
Write-Output "Finalized Excel queue result: $resultPath"
