Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0revision.ps1"
