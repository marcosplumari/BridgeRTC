@echo off
echo =========================================================
echo Registrando SplitPaymentBridge.dll para uso em VB6 (COM)
echo =========================================================
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" "%~dp0bin\Release\SplitPaymentBridge.dll" /codebase /tlb
pause
