@echo off
echo Install as Service...
@nssm install "tsIRCd-1.2" tsIRCd.exe
echo Run the Service...
@nssm start "tsIRCd-1.2"
echo.
echo Done.
pause >nul