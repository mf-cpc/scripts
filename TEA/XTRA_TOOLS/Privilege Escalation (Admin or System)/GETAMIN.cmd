@setlocal enableextensions
@cd /d "%~dp0
@echo off
IF EXIST cd.tmp GOTO one
CD > cd.tmp
:one
if _%1_==_payload_  goto :payload

:getadmin
    echo %~nx0: elevating self
    set vbs=%temp%\getadmin.vbs
    echo Set UAC = CreateObject^("Shell.Application"^)                >> "%vbs%"
    echo UAC.ShellExecute "%~s0", "payload %~sdp0 %*", "", "runas", 1 >> "%vbs%"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
goto :eof

:payload
set /p currentdirectory=<cd.tmp
del cd.tmp
%currentdirectory:~0,2%
cd %currentdirectory:~2,100%
::ENTER YOUR CODE BELOW::   

whoami
pause
cmd



::END OF YOUR CODE::

echo.
echo...Script Complete....
echo.

pause