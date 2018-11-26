@echo off
REM 将参数赋予VBS
cscript /nologo "%~dp0\VBScript\JP_fn_to_U.vbs" %*
echo.
echo 程序运行结束，按任意键结束运行...
pause >nul