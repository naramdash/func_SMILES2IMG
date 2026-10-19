@echo off
chcp 65001 >nul 2>&1
title func_SMILES2IMG - Uninstaller
echo ============================================================================
echo   func_SMILES2IMG (Excel, Word, PowerPoint) - Uninstaller
echo ============================================================================
echo.

if exist "%~dp0uninstall.ps1" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/uninstall.ps1 | iex"
)

echo.
pause
