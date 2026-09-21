@echo off
echo =========================================================
echo Registrando BridgeRTC.dll para uso em VB6 (COM Interop)
echo =========================================================
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "%~dp0bin\Release\BridgeRTC.dll" /codebase /tlb
pause
