@echo off
setlocal
cd /d "%~dp0" || exit /b 1
if not exist "%~dp0t.xlsm" (
  echo Missing prepared workbook: %~dp0t.xlsm. Run prepare-excel-vba.sh before copying the workbook. 1>&2
  exit /b 1
)
if not exist "%~dp0t.xlsm.source.sha256" (
  echo Missing t.xlsm.source.sha256. Run prepare-excel-vba.sh before copying the workbook. 1>&2
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File .\verify-excel-queue-source.ps1 ^
  -SourceDirectory . ^
  -PreparationStamp "t.xlsm.source.sha256"
if errorlevel 1 exit /b 1
powershell -NoProfile -ExecutionPolicy Bypass -File .\skip-if-excel-result-current.ps1 ^
  -SourceDirectory . ^
  -Result "ExcelQueueVerification.result" ^
  -CompletionMarker "QUEUE_COMPLETE" ^
  -SourcePrefix "ExcelQueue"
if errorlevel 2 exit /b 1
if not errorlevel 1 (
  endlocal
  exit /b 0
)
rmdir /s /q "%TEMP%\vba-runner-xl-queue" 2>nul
del /q "%~dp0ExcelQueueVerification.result" 2>nul
rem copy /b /y empty_with_macro.xlsm t.xlsm >nul || exit /b 1
rem call npm run vba-extractor --prefix ../../.. -- import t.xlsm . --yes
rem if errorlevel 1 exit /b 1
del /q t.xlsm.bak 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -File .\run-excel-vba.ps1 ^
  -Workbook "%~dp0t.xlsm" ^
  -Procedure RunExcelQueueVerification ^
  -Module ExcelQueueVerification
if errorlevel 1 exit /b 1
powershell -NoProfile -ExecutionPolicy Bypass -File .\convert-to-utf8.ps1 ^
  "ExcelQueueVerification.result"
if errorlevel 1 exit /b 1
powershell -NoProfile -ExecutionPolicy Bypass -File .\finalize-excel-queue.ps1 ^
  -SourceDirectory . ^
  -Result "ExcelQueueVerification.result"
if errorlevel 1 exit /b 1
endlocal
