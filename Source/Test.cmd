@echo off
setlocal
cd /d "%~dp0"
"%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe" /nologo /target:exe /main:Tests /out:..\Tests.exe /reference:System.dll /reference:System.Core.dll /reference:System.Data.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll /reference:..\ExcelDataReader.dll Program.cs Tests.cs
if errorlevel 1 (
  pause
  exit /b 1
)
"..\Tests.exe" "%~dp0.."
set "TEST_RESULT=%ERRORLEVEL%"
del "..\Tests.exe"
pause
exit /b %TEST_RESULT%
