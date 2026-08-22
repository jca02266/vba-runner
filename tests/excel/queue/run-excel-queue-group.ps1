[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDirectory,
    [Parameter(Mandatory = $true)]
    [string]$Workbook,
    [Parameter(Mandatory = $true)]
    [string]$PreparationStamp,
    [Parameter(Mandatory = $true)]
    [string]$Result,
    [Parameter(Mandatory = $true)]
    [string]$CompletionMarker,
    [Parameter(Mandatory = $true)]
    [string]$SourcePrefix,
    [Parameter(Mandatory = $true)]
    [string]$Procedure,
    [string]$Module,
    [switch]$CleanTemp
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$powershell = (Get-Command powershell.exe -ErrorAction Stop).Source

function Invoke-QueueScript {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [string[]]$Arguments = @()
    )
    & $powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $scriptRoot $Name) @Arguments |
        ForEach-Object { Write-Host $_ }
    return $LASTEXITCODE
}

if (-not (Test-Path -LiteralPath $Workbook -PathType Leaf)) {
    throw "Workbook not found: $Workbook"
}
if (-not (Test-Path -LiteralPath $PreparationStamp -PathType Leaf)) {
    throw "Preparation stamp not found: $PreparationStamp"
}

$verifyCode = Invoke-QueueScript 'verify-excel-queue-source.ps1' @(
    '-SourceDirectory', $SourceDirectory,
    '-PreparationStamp', $PreparationStamp
)
if ($verifyCode -ne 0) {
    exit $verifyCode
}

$skipCode = Invoke-QueueScript 'skip-if-excel-result-current.ps1' @(
    '-SourceDirectory', $SourceDirectory,
    '-Result', $Result,
    '-CompletionMarker', $CompletionMarker,
    '-SourcePrefix', $SourcePrefix
)
if ($skipCode -eq 2) {
    exit 2
}
if ($skipCode -eq 0) {
    exit 0
}
if ($skipCode -ne 1) {
    throw "Unexpected result reuse check exit code: $skipCode"
}

if ($CleanTemp) {
    Remove-Item -LiteralPath (Join-Path $env:TEMP 'vba-runner-xl-queue') -Recurse -Force -ErrorAction SilentlyContinue
}
Remove-Item -LiteralPath $Result -Force -ErrorAction SilentlyContinue

$runArguments = @(
    '-Workbook', $Workbook,
    '-Procedure', $Procedure
)
if ($Module) {
    $runArguments += @('-Module', $Module)
}
$runCode = Invoke-QueueScript 'run-excel-vba.ps1' $runArguments
if ($runCode -ne 0) {
    exit $runCode
}

$convertCode = Invoke-QueueScript 'convert-to-utf8.ps1' @($Result)
if ($convertCode -ne 0) {
    exit $convertCode
}

$finalizeArguments = @(
    '-SourceDirectory', $SourceDirectory,
    '-Result', $Result,
    '-CompletionMarker', $CompletionMarker
)
$finalizeCode = Invoke-QueueScript 'finalize-excel-queue.ps1' $finalizeArguments
if ($finalizeCode -ne 0) {
    exit $finalizeCode
}
exit 0
