[CmdletBinding()]
param(
    [string]$Workbook = (Join-Path $PSScriptRoot 'test.xlsm'),
    [string]$Output = (Join-Path $PSScriptRoot 'ExcelQueueVerification.result')
)

$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null

try {
    if (-not (Test-Path -LiteralPath $Workbook -PathType Leaf)) {
        throw "Workbook not found: $Workbook"
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Open((Resolve-Path -LiteralPath $Workbook).Path)

    # The verification modules are imported into this session only. The
    # workbook is closed with SaveChanges:=False, so test.xlsm is unchanged.
    $components = @($book.VBProject.VBComponents)
    $names = @($components | ForEach-Object { $_.Name })
    $bas = Join-Path $PSScriptRoot 'ExcelQueueVerification.bas'
    $ticket = Join-Path $PSScriptRoot 'ExcelQueueTicket.cls'
    if ($names -notcontains 'ExcelQueueVerification') {
        [void]$book.VBProject.VBComponents.Import((Resolve-Path -LiteralPath $bas).Path)
    }
    $components = @($book.VBProject.VBComponents)
    $names = @($components | ForEach-Object { $_.Name })
    if ($names -notcontains 'ExcelQueueTicket') {
        [void]$book.VBProject.VBComponents.Import((Resolve-Path -LiteralPath $ticket).Path)
    }

    $immediate = $excel.VBE.Windows.Item('Immediate')
    $before = [string]$immediate.Object.Text
    $macro = "'$($book.Name)'!ExcelQueueVerification.RunExcelQueueVerification"
    [void]$excel.Run($macro)
    Start-Sleep -Milliseconds 500
    $after = [string]$immediate.Object.Text
    if ($after.StartsWith($before, [System.StringComparison]::Ordinal)) {
        $result = $after.Substring($before.Length)
    } else {
        $result = $after
    }
    if ([string]::IsNullOrWhiteSpace($result)) {
        throw 'No Debug.Print output was captured from the Immediate window.'
    }

    $outputPath = [System.IO.Path]::GetFullPath($Output)
    $parent = Split-Path -Parent $outputPath
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        [void](New-Item -ItemType Directory -Path $parent)
    }
    [System.IO.File]::WriteAllText($outputPath, $result, (New-Object System.Text.UTF8Encoding($false)))
    Write-Output "Saved Debug.Print output: $outputPath"
}
catch {
    Write-Error ("Excel queue verification failed. Enable 'Trust access to the VBA project object model' " +
        "in Excel Trust Center if VBProject access is denied. Details: " + $_.Exception.Message)
    exit 1
}
finally {
    if ($book) {
        $book.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($book)
    }
    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
}
