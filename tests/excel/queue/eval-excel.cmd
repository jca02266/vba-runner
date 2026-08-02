@echo off
setlocal
cd /d "%~dp0" || exit /b 1
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
