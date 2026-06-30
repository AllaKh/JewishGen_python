@echo off
rem get_machine_id.bat - the RECIPIENT runs this BEFORE installing.
rem It prints this computer's machine id (Windows MachineGuid) and copies it to the
rem clipboard. The recipient sends that line back; you build a package locked to it.
for /f "tokens=3" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Cryptography" /v MachineGuid /reg:64 2^>nul') do set "MID=%%a"
if not defined MID (
  echo Could not read the machine id. Please contact the sender.
  pause
  exit /b 1
)
echo.
echo   Your machine id - send this whole line to the sender:
echo.
echo       %MID%
echo.
<nul set /p "=%MID%"| clip
echo   It was copied to the clipboard - just paste it into your message.
echo.
pause
