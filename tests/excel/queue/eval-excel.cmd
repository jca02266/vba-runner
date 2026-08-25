@echo off
setlocal
cd /d "%~dp0" || exit /b 1
powershell -NoProfile -ExecutionPolicy Bypass -File .\run-excel-queue-group.ps1 ^
  -SourceDirectory . ^
  -Workbook "t.xlsm" ^
  -PreparationStamp "t.xlsm.source.sha256" ^
  -Result "ExcelQueueVerification.result" ^
  -CompletionMarker "QUEUE_COMPLETE" ^
  -SourcePrefix "ExcelQueue" ^
  -Procedure "RunExcelQueueVerification" ^
  -Module "ExcelQueueVerification" ^
  -CleanTemp
if errorlevel 1 exit /b 1
powershell -NoProfile -ExecutionPolicy Bypass -File .\run-excel-isolated-probes.ps1
if errorlevel 1 exit /b 1
endlocal
