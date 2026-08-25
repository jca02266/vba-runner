[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$Workbook,
    [Parameter(Mandatory = $true)] [string]$Module,
    [Parameter(Mandatory = $true)] [string]$Procedure,
    [Parameter(Mandatory = $true)] [string]$Output,
    [Parameter(Mandatory = $true)] [string]$PidFile
)

$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null

function Write-Status {
    param([string[]]$Lines)
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines((Resolve-Path -LiteralPath $Output -ErrorAction SilentlyContinue).Path, $Lines, $utf8)
}

try {
    $outputPath = [System.IO.Path]::GetFullPath($Output)
    $pidPath = [System.IO.Path]::GetFullPath($PidFile)
    $outputDirectory = Split-Path -Parent $outputPath
    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    if (-not ('IsolatedExcelProcess' -as [type])) {
        Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class IsolatedExcelProcess {
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@
    }
    [uint32]$excelPid = 0
    [void][IsolatedExcelProcess]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
    [System.IO.File]::WriteAllText($pidPath, [string]$excelPid)
    $book = $excel.Workbooks.Open((Resolve-Path -LiteralPath $Workbook).Path)
    $qualified = "$Module.$Procedure"
    # Some Excel versions expose imported standard modules through the
    # workbook-level macro name even when the exported Attribute VB_Name is
    # present. Keep the qualified forms first, then use the documented
    # workbook-level fallback before classifying the case as a run failure.
    $macroNames = @(
        "$($book.Name)!$qualified",
        "'$($book.Name)'!$qualified",
        "$($book.Name)!$Procedure",
        "'$($book.Name)'!$Procedure",
        $Procedure
    )
    $runError = $null
    foreach ($macro in $macroNames) {
        try {
            $returnValue = $excel.Run($macro)
            $lines = @('STATUS=ok', "RESULT=$([string]$returnValue)")
            [System.IO.File]::WriteAllLines($outputPath, $lines, (New-Object System.Text.UTF8Encoding($false)))
            exit 0
        } catch {
            $runError = $_
        }
    }
    $message = [string]$runError.Exception.Message
    [System.IO.File]::WriteAllLines($outputPath, @('STATUS=run-error', "ERROR=$message"), (New-Object System.Text.UTF8Encoding($false)))
    exit 2
} catch {
    $message = [string]$_.Exception.Message
    try {
        [System.IO.File]::WriteAllLines(([System.IO.Path]::GetFullPath($Output)), @('STATUS=harness-error', "ERROR=$message"), (New-Object System.Text.UTF8Encoding($false)))
    } catch { }
    exit 3
} finally {
    if ($book) {
        $book.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($book)
    }
    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
}
