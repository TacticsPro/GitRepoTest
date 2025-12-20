@echo off
setlocal

:: Enable Trust Access to VBA Project
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v AccessVBOM /t REG_DWORD /d 1 /f
if %errorlevel% neq 0 (
    echo Error enabling VBA trust access.
) else (
    echo Trust access to VBA project enabled.
)

:: Enable Developer Tab
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Options" /v DeveloperTools /t REG_DWORD /d 1 /f
if %errorlevel% neq 0 (
    echo Error enabling Developer tab.
) else (
    echo Developer tab enabled.
)

:: Enable VBA Macros (Low Security, Not Recommended)
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d 1 /f
if %errorlevel% neq 0 (
    echo Error enabling macros.
) else (
    echo All macros enabled successfully.
)

endlocal

