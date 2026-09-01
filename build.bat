@echo off
setlocal
chcp 65001 > nul

echo ========================================================
echo   pm+helper C# build
echo ========================================================
echo.

set "CSC_PATH=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC_PATH%" set "CSC_PATH=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"

if not exist "%CSC_PATH%" (
    echo [ERROR] csc.exe was not found.
    echo Install or enable .NET Framework 4.x developer tools.
    exit /b 1
)

echo Compiler: %CSC_PATH%
echo Building pm+helper.exe...
echo.

"%CSC_PATH%" ^
  /target:winexe ^
  /out:pm+helper.exe ^
  /optimize+ ^
  /platform:anycpu ^
  /codepage:65001 ^
  /win32icon:PatientHelper.ico ^
  /reference:System.dll,System.Data.dll,System.Drawing.dll,System.Windows.Forms.dll,System.Xml.dll ^
  pm+helper.cs UpdateManager.cs

if errorlevel 1 (
    echo.
    echo ========================================================
    echo   [FAILED] Build failed.
    echo ========================================================
    exit /b 1
)

echo.
echo ========================================================
echo   [OK] pm+helper.exe build completed.
echo ========================================================
exit /b 0
