@echo off
:: ######################################################
:: VERSION 0.9.210908a - CMPFOR2104
:: SCRIPT: HIBERNATE
:: CREATION DATE: 2011-05-26
:: LAST MODIFIED: 2021-09-08
:: AUTHOR: Sgt. Mark SOUTHBY, #53384
:: EMAIL: mark.southby@rcmp-grc.gc.ca
:: ######################################################
:: DESCRIPTION:Attempt to place system into hibernate
:: START OF BATCH
echo Putting system to Hibernate. After system 'shuts down' 
echo please wait a few minutes prior to removing power as writes my still be pending.
:: OS CHECK

:osdetermine
::==========================================================================================
:: --------------------------------DETERMINE Weather to run XP Mode (All USB versions of tools or Win 7 Mode.
:: --------------------------------Windows 7 / vista etc. Require the default application to run, this is a protection feature
:: --------------------------------Of that OS to prevent malicous programs running under the same name. These commands can be extracted from
:: --------------------------------Windows/System32/ directory and hashed/compared with known version to determine if they are OS versions.

ver | find "2003" > nul
if %ERRORLEVEL% == 0 goto ver_xp
ver | find "XP" > nul
if %ERRORLEVEL% == 0 goto ver_xp
ver | find "2000" > nul
if %ERRORLEVEL% == 0 goto ver_xp
ver | find "NT" > nul
if %ERRORLEVEL% == 0 goto ver_xp
goto Win7

:xp
::RunDll32.exe shell32.dll,Control_RunDLL hotplug.dll
%windir%\system32\rundll32.exe powrprof.dll,SetSuspendState hibernate

:Win7
powercfg.exe /hibernate on
shutdown /h
