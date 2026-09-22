@echo off
title Cai Dat ExcelSupport Add-In
dotnet build -c release
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-AddIn.ps1"
