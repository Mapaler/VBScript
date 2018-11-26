@echo off
REM 将参数赋予VBS
cscript /nologo "%~dp0\VBScript\solid_7z_to_jpg_fast.vbs" %*
echo.
echo 程序运行结束，按任意键结束运行...
pause >nul