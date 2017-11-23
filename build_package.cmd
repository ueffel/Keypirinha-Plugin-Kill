@echo off
SET SEVENZIP=

where 7z > nul 2>&1
if not errorlevel 1 (
    set SEVENZIP=7z
    goto pack
)

where 7za > nul 2>&1
if not errorlevel 1 (
    set SEVENZIP=7za
    goto pack
)

if exist "c:\Program Files (x86)\7-Zip\7z.exe" (
    SET "SEVENZIP=c:\Program Files (x86)\7-Zip\7z.exe"
    goto pack
)

if exist "c:\Program Files (x86)\7-Zip\7za.exe" (
    SET "SEVENZIP=c:\Program Files (x86)\7-Zip\7za.exe"
    goto pack
)

if exist "c:\Program Files\7-Zip\7z.exe" (
    SET "SEVENZIP=c:\Program Files\7-Zip\7z.exe"
    goto pack
)

if exist "c:\Program Files\7-Zip\7za.exe" (
    SET "SEVENZIP=c:\Program Files\7-Zip\7za.exe"
    goto pack
)

if NOT DEFINED SEVENZIP (
    echo 7zip not found
    exit /b 1
)

:pack
if exist Kill.keypirinha-package (
    del Kill.keypirinha-package
)
echo Using "%SEVENZIP%" to pack
"%SEVENZIP%" a -mx9 -tzip Kill.keypirinha-package  -x!%~nx0 -xr!.git *
