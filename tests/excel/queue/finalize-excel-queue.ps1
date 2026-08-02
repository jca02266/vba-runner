[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDirectory,
    [Parameter(Mandatory = $true)]
    [string]$Result
)

$ErrorActionPreference = 'Stop'
$sourceDirectoryPath = (Resolve-Path -LiteralPath $SourceDirectory -ErrorAction Stop).Path
$resultPath = (Resolve-Path -LiteralPath $Result -ErrorAction Stop).Path
$utf8 = New-Object System.Text.UTF8Encoding($false)
$resultText = [System.IO.File]::ReadAllText($resultPath, $utf8)

if ($resultText -notmatch '(?m)^QUEUE_COMPLETE=True\s*$') {
    throw "Excel queue did not reach its completion marker: $resultPath"
}

$sourceFiles = Get-ChildItem -LiteralPath $sourceDirectoryPath -File |
    Where-Object { $_.Extension -match '^\.(bas|cls|frm)$' } |
    Sort-Object -Property Name
if ($sourceFiles.Count -eq 0) {
    throw "No VBA source files found: $sourceDirectoryPath"
}
$manifest = New-Object System.Text.StringBuilder
foreach ($sourceFile in $sourceFiles) {
    $sourceText = [System.IO.File]::ReadAllText($sourceFile.FullName, $utf8).Replace("`r`n", "`n")
    [void]$manifest.Append($sourceFile.Name).Append("`n").Append($sourceText).Append("`n")
}
$sha256 = [System.Security.Cryptography.SHA256]::Create()
try {
    $hashBytes = $sha256.ComputeHash($utf8.GetBytes($manifest.ToString()))
} finally {
    $sha256.Dispose()
}
$hash = ([System.BitConverter]::ToString($hashBytes)).Replace('-', '').ToLowerInvariant()
$resultText = $resultText.TrimEnd("`r", "`n") + "`nQUEUE_SOURCE_SHA256=$hash`n"
[System.IO.File]::WriteAllText($resultPath, $resultText, $utf8)
Write-Output "Finalized Excel queue result: $resultPath"
