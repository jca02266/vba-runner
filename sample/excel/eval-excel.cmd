@echo off
setlocal
cd /d "%~dp0" || exit /b 1
rmdir /s /q "%TEMP%\vba-runner-xl-queue" 2>nul
copy /b /y empty_with_macro.xlsm t.xlsm >nul || exit /b 1
call npm run vba-extractor --prefix ../.. -- import t.xlsm . --yes
if errorlevel 1 exit /b 1
del /q t.xlsm.bak 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -File .\run-excel-queue.ps1 -Workbook "%~dp0t.xlsm" -Procedure RunExcelQueueVerification -Module ExcelQueueVerification -Output "%~dp0ExcelQueueVerification.result"
endlocal
