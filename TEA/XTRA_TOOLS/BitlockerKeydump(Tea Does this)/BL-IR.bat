@echo off
title BLIR
cls
:: ######################################################
:: VERSION 1.0.131008 - October 08 2013
:: SCRIPT: Bit Locker Incident Response (BLIR)
:: CREATION DATE: 2013-10-08
:: LAST MODIFIED: 2013-10-09
:: AUTHOR: Cst. Mark SOUTHBY
:: EMAIL: mark@southby.ca
:: ######################################################
:: #### Authored by Cst. SOUTHBY, for personal use. Not sanctioned or approved by any organization or entity. Use at your own risk.
:: ---------------------------------------------- 
:: #### For Law enforcement use only. Can be used and modified without permission for Law Enforcement purposes.
:: START OF BATCH

::Run only if admin
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"&&( goto start)
goto noadmin

:start
:: Set date time into varible so no output files are overwritten
::replace / with -
set SAVEDATE=%DATE:/=-%
set SAVEDATE=%SAVEDATE: =%
::remove any date words
set SAVEDATE=%SAVEDATE:Mon=%
set SAVEDATE=%SAVEDATE:Tue=%
set SAVEDATE=%SAVEDATE:Wed=%
set SAVEDATE=%SAVEDATE:Thu=%
set SAVEDATE=%SAVEDATE:Fri=%
set SAVEDATE=%SAVEDATE:Sat=%
set SAVEDATE=%SAVEDATE:Sun=%
::replace : with .
set SAVETIME=%TIME::=.%
::Remove fractions of seconds
set SAVETIME=%SAVETIME:~0,8%
::set output filename with date and time
set OUTPUT=bitlocker_status@%SAVEDATE%_%SAVETIME%.txt

:: Provide default varible for %drive%
set drive=.
:: find bitlocker status ( Protection on or off)
echo Searching for Bitlocker...
!cscript !manage-bde.wsf -status >%OUTPUT%
find /c /i "Protection On" %OUTPUT% > NUL
if %ERRORLEVEL% == 0 goto found
goto notfound

:found
color 0c
echo.
echo.
echo ENCRYPTED VOLUME FOUND, Please check results!
echo.
echo Look under "Conversion Status" for "Fully Encrypted"
echo Please note down Volume letter.
pause
color
type %OUTPUT%|more
pause
color
goto step2

:notfound
echo No BITLOCKER found. Please verify output.
pause
type %OUTPUT%|more
goto exit
pause

:step2
echo What is the Drive letter of the Volume (ie C:)
set /p drive=Type letter only or . if no volume:
if %drive% == . goto fail
echo Generating a recovery key. (Bottom "Password" will access bitlocker)
!cscript !manage-bde.wsf -protectors -get %drive%: > %drive%_Bitlocker_recovery_key.txt
type %drive%_Bitlocker_recovery_key.txt
pause
goto exit

:fail
cls
echo Encrypted Volume not assigned a drive letter.
echo Please turn off Bitlocker for volume using GUI 
echo.
echo [Press start, type bit, select "Manage Bitlocker"]
echo.
goto exit
:: %windir%\system32\Bitlockerwizardelev.exe \ t

:: Disable Protection (Does not DECRYPT, stores key on disk so it can be decrypted when accessed with Windows 7 or Vista)
::cscript manage-bde.wsf -protectors -disable c:
::pause
:noadmin
cls
color
echo Not run as an administrator account. 
echo.
echo Please right click and try run command line as an administrator 
echo (or press start type 'cmd' then press control-shift-enter) 
echo If you can not run as an administrator, try running GetSystem.cmd
pause
goto exit

:exit
color