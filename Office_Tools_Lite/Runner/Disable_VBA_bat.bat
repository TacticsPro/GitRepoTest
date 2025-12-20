@echo off
setlocal

:: Disable Trust Access to VBA Project
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v AccessVBOM /t REG_DWORD /d 0 /f
if %errorlevel% neq 0 (
    echo Error disabling VBA trust access.
) else (
    echo Trust access to VBA project disabled.
)

:: Disable Developer Tab
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Options" /v DeveloperTools /t REG_DWORD /d 0 /f
if %errorlevel% neq 0 (
    echo Error disabling Developer tab.
) else (
    echo Developer tab disabled.
)

:: Disable VBA Macros (Set to high security, prompt for macros)
reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v VBAWarnings /t REG_DWORD /d 2 /f
if %errorlevel% neq 0 (
    echo Error setting macro security level.
) else (
    echo Macro execution set to prompt (macros disabled by default).
)

endlocal