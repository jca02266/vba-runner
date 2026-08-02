[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Workbook,
    [Parameter(Mandatory = $true)]
    [string]$Procedure,
    [string]$Module,
    [string]$Output
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

    $immediate = $excel.VBE.Windows.Item('Immediate')
    $before = [string]$immediate.Object.Text
    $qualifiedProcedure = if ($Module) { "$Module.$Procedure" } else { $Procedure }
    $macro = "'$($book.Name)'!$qualifiedProcedure"
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

    if (-not $Output) {
        $Output = Join-Path (Split-Path -Parent (Resolve-Path -LiteralPath $Workbook).Path) `
            "$([System.IO.Path]::GetFileNameWithoutExtension($Workbook)).result"
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
    Write-Error ("Excel macro execution failed. Ensure the workbook already contains the requested " +
        "module/procedure and that the Immediate window is available. Details: " + $_.Exception.Message)
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
