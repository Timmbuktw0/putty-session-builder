@echo off
setlocal
cd /d "%~dp0"
set "COMPILER=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%COMPILER%" (
  echo .NET Framework Compiler nicht gefunden.
  pause
  exit /b 1
)
"%COMPILER%" /nologo /target:winexe /platform:anycpu /optimize+ /win32manifest:app.manifest /out:..\PuTTYSessionBuilder.exe /reference:System.dll /reference:System.Core.dll /reference:System.Data.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll /reference:..\ExcelDataReader.dll Program.cs
if errorlevel 1 (
  pause
  exit /b 1
)
echo Programm erfolgreich gebaut.
pause
