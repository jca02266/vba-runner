<#
.SYNOPSIS
    Build a new .xlsm file by importing all .bas files under a source directory.

.DESCRIPTION
    vba-extractor's "import" command requires an existing .xlsm file, so it can't
    create one from scratch. This script bootstraps the first .xlsm by driving
    Excel via COM automation: create a new workbook, import each .bas module,
    then save as .xlsm.

    Requires Windows + a locally installed Excel (COM automation). Excel's
    "Trust access to the VBA project object model" setting (AccessVBOM) must be
    enabled beforehand, e.g.:

        New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Excel\Security" `
            -Name "AccessVBOM" -Value 1 -PropertyType DWord -Force

    If a .bas file doesn't start with an "Attribute VB_Name = ..." line, Excel
    assigns a default name like Module1/Module2 instead of the intended module
    name. To keep .bas source files free of this Excel-only detail, this script
    injects the Attribute line (derived from the file name) into a temporary
    copy right before importing.

    This script only imports standard modules (.bas). Class modules, forms, and
    document modules (ThisWorkbook/Sheet1 etc., .cls) are not handled here -
    after running this script, use `vba-extractor import <OutputPath> <SourceDir>`
    to sync those (and re-sync the .bas files) into the newly created workbook.

.PARAMETER SourceDir
    Directory containing the .bas files to import.

.PARAMETER OutputPath
    Path of the .xlsm file to create.

.EXAMPLE
    powershell -File Build-Xlsm.ps1 -SourceDir src\vba -OutputPath MyBook.xlsm
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDir,
    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$xlOpenXMLWorkbookMacroEnabled = 52

if (-not (Test-Path $SourceDir)) {
    throw "SourceDir not found: $SourceDir"
}
if (Test-Path $OutputPath) {
    Remove-Item $OutputPath -Force
}

$basFiles = Get-ChildItem -Path $SourceDir -Filter "*.bas"
if ($basFiles.Count -eq 0) {
    throw "No .bas files found in: $SourceDir"
}

$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("vba-import-" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Add()

    foreach ($file in $basFiles) {
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $content = Get-Content -Path $file.FullName -Raw -Encoding UTF8

        if ($content -notmatch '^\s*Attribute\s+VB_Name\s*=') {
            $content = "Attribute VB_Name = `"$moduleName`"`r`n$content"
        }

        # VBA module import expects the system ANSI codepage, not UTF-8.
        $tempFile = "$tempDir\$($file.Name)"
        Set-Content -Path $tempFile -Value $content -Encoding Default -NoNewline

        Write-Host "Importing: $($file.Name) (as $moduleName)"
        $workbook.VBProject.VBComponents.Import($tempFile) | Out-Null
    }

    $workbook.SaveAs($OutputPath, $xlOpenXMLWorkbookMacroEnabled)
    Write-Host "Saved: $OutputPath"
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}
