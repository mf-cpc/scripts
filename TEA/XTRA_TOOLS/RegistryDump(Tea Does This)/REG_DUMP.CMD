@echo off
:: SCRIPT: Registry Dump
:: CREATION DATE: 2021-04-08
:: LAST MODIFIED: 2021-04-08
:: AUTHOR: Mark SOUTHBY
:: EMAIL: mark@southby.ca
:: ######################################################
:: This tool is designed to be used with TEA
:: Steps:
:: Launch a command prompt as Administrator (windows key, cmd, control+shift+enter)
:: Browse to TEA drive and folder
:: Run GETSYSTEM.CMD
:: Browse to TEA drive and folder again
:: Run REG_DUMP.CMD
:: ######################################################
:: Known Bugs 
:: - If there was a second instance of tea run with the same CFG, this will put the registry in the first folder. 
:: Running it again, even on a new computer will over write it. 
:: BREW is supposed to be run between each exibit from a forensic workstation
:: ######################################################
TITLE ***TEA Registry Dumper***
CLS
ECHO This needs to be run as system.
ECHO You are logged in as:
ECHO.
ECHO.
whoami
ECHO.
ECHO.
ECHO This tool is designed to be used with TEA
ECHO Steps:
ECHO 1) Launch a command prompt as Administrator 
echo (windows key, cmd, control+shift+enter)
ECHO 2) Browse to TEA drive and folder
ECHO 3) Run GETSYSTEM.CMD
ECHO 4) Browse to TEA drive and folder again
ECHO 5) Run REG_DUMP.CMD
ECHO.
pause
IF EXIST tea.cfg goto found
ECHO No TEA.CFG found, exporting to this directory:
cd |findstr :
::powershell reg save HKLM\SAM SAM
::powershell reg save HKLM\SECURITY SECURITY
::powershell reg save HKLM\SOFTWARE SOFTWARE
::powershell reg save HKLM\SYSTEM SYSTEM
GOTO exit
:FOUND
FOR /F "tokens=2 delims==" %%a IN ('find "filenum" ^<tea.cfg') DO SET filenum=%%a
FOR /F "tokens=2 delims==" %%b IN ('find "exnum" ^<tea.cfg') DO SET exnum=%%b
ECHO tea.cfg found. Registry being dumped to %filenum%-%exnum%\REG\
mkdir %filenum%-%exnum%\REG\
IF EXIST C:\Windows\System32\config\SAM goto Cdrive
:Cdrive
extents-win64.exe  C:\Windows\System32\config\SAM %filenum%-%exnum%\REG\SAM
extents-win64.exe  C:\Windows\System32\config\SECURITY %filenum%-%exnum%\REG\SECURITY
extents-win64.exe  C:\Windows\System32\config\SOFTWARE %filenum%-%exnum%\REG\SOFTWARE
extents-win64.exe  C:\Windows\System32\config\SYSTEM %filenum%-%exnum%\REG\SYSTEM
::powershell reg save HKLM\SAM %filenum%-%exnum%\REG\SAM
::powershell reg save HKLM\SECURITY %filenum%-%exnum%\REG\SECURITY
::powershell reg save HKLM\SOFTWARE %filenum%-%exnum%\REG\SOFTWARE
::powershell reg save HKLM\SYSTEM %filenum%-%exnum%\REG\SYSTEM
::powershell reg save HKCU %filenum%-%exnum%\REG\NTUser.DAT

:NTUSER
ECHO.
ECHO Listing NTUser.dat files ...
ECHO.
dir /a/b/s c:\users\*NTUser.dat
echo.
echo WHEN DONE, PRESS ENTER
set path1=EXIT
Set /p path1=Select path to NTUSER.DAT to copy and right click, then right click to paste here:
if %path1% == EXIT goto EXIT
:run1
mkdir %filenum%-%exnum%\REG\%path1:~9,-10%
extents-win64.exe %path1% %filenum%-%exnum%\REG\%path1:~9,100%
goto NTUSER



:exit