[CmdletBinding()]
param(
    [int]$TimeoutSeconds = 120
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$isolated = Join-Path $scriptRoot 'isolated'
$workbooks = Join-Path $isolated 'workbooks'
$manifestPath = Join-Path $workbooks 'manifest.json'
$stampPath = Join-Path $workbooks 'source.sha256'
$resultPath = Join-Path $scriptRoot 'IsolatedCompileProbe.result'
$runner = Join-Path $scriptRoot 'run-excel-compile-probe.ps1'

if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Isolated probe manifest not found: $manifestPath" }
if (-not (Test-Path -LiteralPath $stampPath -PathType Leaf)) { throw "Isolated probe preparation stamp not found: $stampPath" }
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json

function Get-SourceHash {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $builder = New-Object System.Text.StringBuilder
        $sources = Get-ChildItem -LiteralPath $isolated -File |
            Where-Object { $_.Extension -match '^\.(bas|cls|frm)$' } |
            Sort-Object -Property Name
        foreach ($source in $sources) {
            $text = [System.IO.File]::ReadAllText($source.FullName, (New-Object System.Text.UTF8Encoding($false))).Replace("`r`n", "`n")
            [void]$builder.Append($source.Name).Append("`n").Append($text).Append("`n")
        }
        $bytes = $sha.ComputeHash((New-Object System.Text.UTF8Encoding($false)).GetBytes($builder.ToString()))
        return ([System.BitConverter]::ToString($bytes)).Replace('-', '').ToLowerInvariant()
    } finally { $sha.Dispose() }
}

$currentHash = Get-SourceHash
$preparedHash = (Get-Content -LiteralPath $stampPath -Raw).Trim().ToLowerInvariant()
if ($currentHash -ne $preparedHash -or $currentHash -ne ([string]$manifest.sourceHash).ToLowerInvariant()) {
    throw "Isolated probe source is stale; rerun prepare-excel-vba.sh (prepared=$preparedHash current=$currentHash)"
}
foreach ($case in $manifest.cases) {
    $workbook = Join-Path $isolated $case.workbook
    if (-not (Test-Path -LiteralPath $workbook -PathType Leaf)) { throw "Isolated probe workbook not found: $workbook" }
}

$existing = if (Test-Path -LiteralPath $resultPath -PathType Leaf) { Get-Content -LiteralPath $resultPath -Raw } else { '' }
$expectedIds = @($manifest.cases | Where-Object { $_.id -ne 'CONTROL' } | ForEach-Object { [string]$_.id })
$alreadyComplete = $existing -match '(?m)^ISOLATED_COMPILE_PROBE_COMPLETE=True\s*$' -and
    $existing -match "(?m)^QUEUE_SOURCE_SHA256=$([regex]::Escape($currentHash))\s*$"
$allExisting = $alreadyComplete
foreach ($id in $expectedIds) {
    if ($existing -notmatch "(?m)^$([regex]::Escape($id))\b") { $allExisting = $false; break }
}
if ($allExisting) { Write-Output "Reusing current isolated compile result: $resultPath"; exit 0 }

$temp = Join-Path $env:TEMP ("vba-runner-isolated-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp -Force | Out-Null
$lines = New-Object System.Collections.Generic.List[string]
$completed = $true
try {
    foreach ($case in $manifest.cases) {
        $workbook = Join-Path $isolated $case.workbook
        $caseOutput = Join-Path $temp ($case.source + '.out')
        $pidFile = Join-Path $temp ($case.source + '.pid')
        Remove-Item -LiteralPath $caseOutput, $pidFile -Force -ErrorAction SilentlyContinue
        $argumentList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $runner,
            '-Workbook', $workbook, '-Module', [string]$case.module,
            '-Procedure', [string]$case.procedure, '-Output', $caseOutput, '-PidFile', $pidFile)
        $process = Start-Process -FilePath 'powershell.exe' -ArgumentList $argumentList -PassThru -WindowStyle Hidden
        $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
        while (-not $process.HasExited -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 200; $process.Refresh() }
        if (-not $process.HasExited) {
            $completed = $false
            if (Test-Path -LiteralPath $pidFile) {
                $excelPid = [int](Get-Content -LiteralPath $pidFile -Raw)
                if ($excelPid -gt 0) { Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue }
            }
            Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
            $lines.Add("HARNESS_UNVERIFIED PROBE=$($case.id) CASE=$($case.caseName) TIMEOUT=True")
            continue
        }
        $output = if (Test-Path -LiteralPath $caseOutput) { Get-Content -LiteralPath $caseOutput -Raw } else { '' }
        $status = ($output -split "`r?`n" | Where-Object { $_ -match '^STATUS=' } | Select-Object -First 1)
        if ($status -eq 'STATUS=ok') {
            $line = ($output -split "`r?`n" | Where-Object { $_ -match '^RESULT=' } | Select-Object -First 1)
            if ($line) { $lines.Add($line.Substring(7)); continue }
            $completed = $false
            $lines.Add("HARNESS_UNVERIFIED PROBE=$($case.id) CASE=$($case.caseName) MISSING_RESULT=True")
        } elseif ($status -eq 'STATUS=run-error' -and $case.id -ne 'CONTROL') {
            $error = ($output -split "`r?`n" | Where-Object { $_ -match '^ERROR=' } | Select-Object -First 1)
            $message = if ($error) { $error.Substring(6) } else { 'Application.Run failed' }
            $message = $message.Replace("`r", ' ').Replace("`n", ' ')
            $lines.Add("$($case.id) CASE=$($case.caseName) COMPILE=ERROR MESSAGE=[$message]")
        } else {
            throw "Control or harness failure for $($case.source): $output"
        }
    }
    $lines.Add("QUEUE_SOURCE_SHA256=$currentHash")
    $lines.Add("ISOLATED_COMPILE_PROBE_COMPLETE=$([string]$completed)")
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines($resultPath, $lines, $utf8)
    if (-not $completed) { exit 2 }
    Write-Output "Wrote isolated compile result: $resultPath"
} finally {
    Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue
}
