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
  -Result "FormatMatrix.result" ^
  -CompletionMarker "FORMAT_MATRIX_COMPLETE" ^
  -SourcePrefix "FormatMatrix"
if errorlevel 2 exit /b 1
if errorlevel 1 goto run_format_matrix
goto format_matrix_done

:run_format_matrix
del /q "%~dp0FormatMatrix.result" 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -File .\run-excel-vba.ps1 ^
  -Workbook "%~dp0t.xlsm" ^
  -Procedure RunFormatMatrixVerification ^
  -Module FormatMatrixVerification
if errorlevel 1 exit /b 1
powershell -NoProfile -ExecutionPolicy Bypass -File .\convert-to-utf8.ps1 ^
  "FormatMatrix.result"
if errorlevel 1 exit /b 1
powershell -NoProfile -ExecutionPolicy Bypass -File .\finalize-excel-queue.ps1 ^
  -SourceDirectory . ^
  -Result "FormatMatrix.result" ^
  -CompletionMarker "FORMAT_MATRIX_COMPLETE"
if errorlevel 1 exit /b 1

:format_matrix_done
powershell -NoProfile -ExecutionPolicy Bypass -File .\skip-if-excel-result-current.ps1 ^
  -SourceDirectory . ^
  -Result "RadixMatrix.result" ^
  -CompletionMarker "RADIX_MATRIX_COMPLETE" ^
  -SourcePrefix "RadixMatrix"
if errorlevel 2 exit /b 1
if errorlevel 1 goto run_radix_matrix
goto radix_matrix_done

:run_radix_matrix
del /q "%~dp0RadixMatrix.result" 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -File .\run-excel-vba.ps1 ^
  -Workbook "%~dp0t.xlsm" ^
  -Procedure RunRadixMatrixVerification ^
  -Module RadixMatrixVerification
if errorlevel 1 exit /b 1
powershell -NoProfile -ExecutionPolicy Bypass -File .\convert-to-utf8.ps1 ^
  "RadixMatrix.result"
if errorlevel 1 exit /b 1
powershell -NoProfile -ExecutionPolicy Bypass -File .\finalize-excel-queue.ps1 ^
  -SourceDirectory . ^
  -Result "RadixMatrix.result" ^
  -CompletionMarker "RADIX_MATRIX_COMPLETE"
if errorlevel 1 exit /b 1

:radix_matrix_done
endlocal
