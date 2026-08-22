@echo off
setlocal
cd /d "%~dp0" || exit /b 1
powershell -NoProfile -ExecutionPolicy Bypass -File .\run-excel-queue-group.ps1 ^
  -SourceDirectory . ^
  -Workbook "t.xlsm" ^
  -PreparationStamp "t.xlsm.source.sha256" ^
  -Result "FormatMatrix.result" ^
  -CompletionMarker "FORMAT_MATRIX_COMPLETE" ^
  -SourcePrefix "FormatMatrix" ^
  -Procedure "RunFormatMatrixVerification" ^
  -Module "FormatMatrixVerification"
if errorlevel 1 exit /b 1
powershell -NoProfile -ExecutionPolicy Bypass -File .\run-excel-queue-group.ps1 ^
  -SourceDirectory . ^
  -Workbook "t.xlsm" ^
  -PreparationStamp "t.xlsm.source.sha256" ^
  -Result "RadixMatrix.result" ^
  -CompletionMarker "RADIX_MATRIX_COMPLETE" ^
  -SourcePrefix "RadixMatrix" ^
  -Procedure "RunRadixMatrixVerification" ^
  -Module "RadixMatrixVerification"
if errorlevel 1 exit /b 1
endlocal
