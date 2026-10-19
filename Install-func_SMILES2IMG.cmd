@echo off
chcp 65001 >nul 2>&1
title func_SMILES2IMG (Excel, Word, PowerPoint) - One-Click Installer
echo ============================================================================
echo   func_SMILES2IMG (Excel, Word, PowerPoint) - One-Click Installer
echo   - Supports Microsoft Excel, Word, and PowerPoint (32-bit ^& 64-bit)
echo   - No Administrator rights required (installs to current user profile)
echo   - 100%% Offline chemical rendering ^& property calculation after install
echo ============================================================================
echo.

if exist "%~dp0install.ps1" (
    echo [+] Found local install.ps1 in current folder. Starting offline/local install...
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
) else (
    echo [+] Downloading and installing the latest release from GitHub...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; irm https://raw.githubusercontent.com/naramdash/func_SMILES2IMG/main/install.ps1 | iex"
)

echo.
echo ============================================================================
echo   Done! You can now open Excel, Word, or PowerPoint.
echo   Try typing =NAME2IMG("Aspirin") in Excel or pressing Ctrl+Shift+K!
echo ============================================================================
pause
