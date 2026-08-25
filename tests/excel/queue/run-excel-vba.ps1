[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Workbook,
    [Parameter(Mandatory = $true)]
    [string]$Procedure,
    [string]$Module
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
    # msoAutomationSecurityLow (1) applies to workbooks opened by this
    # automation instance.  Without setting it before Open(), Excel may load
    # the workbook with macros disabled and Application.Run then reports the
    # misleading "macros may be disabled" error.
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open((Resolve-Path -LiteralPath $Workbook).Path)

    $qualifiedProcedure = if ($Module) { "$Module.$Procedure" } else { $Procedure }
    # Application.Run accepts <workbook>!<module>.<procedure>.  Excel's
    # parser is inconsistent about single-quoting a workbook name when it is
    # called through COM, so try the documented bare form first and the quoted
    # form only for the specific "cannot run macro" lookup failure.
    $macroCandidates = @(
        "$($book.Name)!$qualifiedProcedure",
        "'$($book.Name)'!$qualifiedProcedure",
        "$($book.Name)!$Procedure",
        "'$($book.Name)'!$Procedure",
        $Procedure
    )
    $runSucceeded = $false
    $runError = $null
    foreach ($macro in $macroCandidates) {
        Write-Verbose "Running Excel macro: $macro"
        try {
            [void]$excel.Run($macro)
            $runSucceeded = $true
            break
        } catch {
            $runError = $_
            $message = [string]$_.Exception.Message
            if ($message -notmatch 'macro|マクロ|使用できません|無効') {
                throw
            }
        }
    }
    if (-not $runSucceeded) {
        throw "Application.Run could not resolve '$qualifiedProcedure' in workbook '$($book.Name)'. Tried: $($macroCandidates -join ', '). Last error: $($runError.Exception.Message)"
    }
    Write-Output "Excel macro completed: $qualifiedProcedure"
}
catch {
    Write-Error ("Excel macro execution failed. Ensure the workbook contains the requested " +
        "module/procedure, and macros are allowed for this automation instance. A file downloaded " +
        "from the Internet may also need to be unblocked " +
        "or opened from a Trusted Location. Details: " + $_.Exception.Message)
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
