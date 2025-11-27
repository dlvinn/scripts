@echo off
setlocal enabledelayedexpansion

echo ========================================
echo ODT to DOCX Batch Converter
echo ========================================
echo.

REM Try to find conversion tool (LibreOffice or WPS)
set "CONVERTER="
set "CONVERT_CMD="

REM Check for LibreOffice
if exist "C:\Program Files\LibreOffice\program\soffice.exe" (
    set "CONVERTER=C:\Program Files\LibreOffice\program\soffice.exe"
    set "CONVERT_CMD=libreoffice"
    echo Found: LibreOffice
    goto :found
)

REM Check for LibreOffice (x86)
if exist "C:\Program Files (x86)\LibreOffice\program\soffice.exe" (
    set "CONVERTER=C:\Program Files (x86)\LibreOffice\program\soffice.exe"
    set "CONVERT_CMD=libreoffice"
    echo Found: LibreOffice (x86)
    goto :found
)

REM Check for WPS Office by trying to create COM object
echo Checking for WPS Office...
cscript //nologo //E:vbscript "%~dp0check_wps.vbs" >nul 2>&1
if %errorlevel%==0 (
    set "CONVERT_CMD=wps"
    echo Found: WPS Office
    goto :found
)

:found
if not defined CONVERT_CMD (
    echo ERROR: No converter found!
    echo.
    echo Please install one of the following:
    echo   - LibreOffice: https://www.libreoffice.org/download/
    echo   - WPS Office: https://www.wps.com/
    echo.
    pause
    exit /b 1
)

if "!CONVERT_CMD!"=="libreoffice" (
    echo Converter: LibreOffice
    echo Using: %CONVERTER%
    echo.
    echo Scanning for .odt files...
    echo.

    set "count=0"
    set "success=0"
    set "failed=0"

    for /R "%cd%" %%F in (*.odt) do (
        set /a count+=1
        echo [!count!] Converting: %%~nxF
        echo     Location: %%~dpF

        REM LibreOffice conversion
        "!CONVERTER!" --headless --convert-to docx --outdir "%%~dpF" "%%F" >nul 2>&1

        if exist "%%~dpF%%~nF.docx" (
            echo     Status: SUCCESS
            set /a success+=1
        ) else (
            echo     Status: FAILED
            set /a failed+=1
        )
        echo.
    )

    echo ========================================
    echo CONVERSION SUMMARY
    echo ========================================
    if !count!==0 (
        echo No .odt files found in current directory or subdirectories.
    ) else (
        echo Total files found:      !count!
        echo Successfully converted: !success!
        echo Failed:                 !failed!
    )
    echo ========================================

) else if "!CONVERT_CMD!"=="wps" (
    echo Converter: WPS Office
    echo Using VBScript automation
    echo.

    REM Use VBScript for WPS conversion
    cscript //nologo "%~dp0convert_wps.vbs"
)

echo.
pause
exit /b 0
