@echo off
:: ######################################################
:: VERSION 1.0.211213a - CMPFOR Edition
:: SCRIPT: TEA TOOL VERIFICATION
:: CREATION DATE: 2011-02-01
:: LAST MODIFIED: 2021-12-13
:: AUTHOR: Mark SOUTHBY, #53384
:: EMAIL: mark.southby@rcmp-grc.gc.ca
:: ######################################################
:: #### For Law enforcement use only. Can be used and modified without permission for Law Enforcement purposes.
title TEA TOOL VERIFICATION
color 02
echo TEA TOOL VERIFICATION
:: HASH AND VERIFY TOOLS

:: clear varibles
set !!arp=0
set !!at=0
set !!crypthunter=0
set !!EDD=0
set !!FDPro=0
::set !!getip=0
set !!curl=0
set !!getmac=0
set !!ipconfig=0
set !!md5sum=0
set !!net=0
set !!netstat=0
set !!nbtstat=0
set !!now=0
set !!openfiles=0
set !!openports=0
set !!psfile=0
set !!Psinfo=0
set !!pslist=0
set !!psloggedon=0
set !!route=0
set !!Sniffer=0
set !!strings=0
set !!systeminfo=0
set !!WHOAMI=0
set !!winen=0
set !!winen_=0
set !!winen64=0
set !!winen64_=0
::set !!driveid=0
set !!psexec=0
set !!win32dd=0
set !!win32ddsys=0
set !!win64dd=0
set !!win64ddsys=0
set !!winpmem=0

:: Hash tools
!!md5sum !!arp.exe > !!arp.md5
!!md5sum !!at.exe > !!at.md5
!!md5sum !!crypthunter.exe > !!crypthunter.md5
!!md5sum !!EDDv301.exe >!!EDD.md5
!!md5sum !!FDPro.exe >!!FDPro.md5
::!!md5sum !!getip.exe >!!getip.md5
!!md5sum !!curl.exe >!!curl.md5
!!md5sum !!getmac.exe >!!getmac.md5
!!md5sum !!ipconfig.exe >!!ipconfig.md5
!!md5sum !!md5sum.exe >!!md5sum.md5
!!md5sum !!net.exe > !!net.md5
!!md5sum !!netstat.exe > !!netstat.md5
!!md5sum !!nbtstat.exe >!!nbtstat.md5
!!md5sum !!now.exe >!!now.md5
!!md5sum !!openfiles.exe >!!openfiles.md5
!!md5sum !!openports.exe >!!openports.md5
!!md5sum !!psfile.exe >!!psfile.md5
!!md5sum !!Psinfo.exe >!!Psinfo.md5
!!md5sum !!pslist.exe >!!pslist.md5
!!md5sum !!psloggedon.exe >!!psloggedon.md5
!!md5sum !!route.exe >!!route.md5
!!md5sum !!Sniffer.exe >!!Sniffer.md5
!!md5sum !!strings.exe >!!strings.md5
!!md5sum !!systeminfo.exe >!!systeminfo.md5
!!md5sum !!WHOAMI.EXE >!!WHOAMI.md5
!!md5sum !!winen.exe >!!winen.md5
!!md5sum winen_.sys >!!winen_.md5
!!md5sum !!winen64.exe >!!winen64.md5
!!md5sum winen64_.sys >!!winen64_.md5
::!!md5sum !!driveid.exe >!!driveid.md5
!!md5sum !!psexec.exe >!!psexec.md5
!!md5sum !!win32dd.exe >!!win32dd.md5
!!md5sum win32dd.sys >!!win32ddsys.md5
!!md5sum !!win64dd.exe >!!win64dd.md5
!!md5sum win64dd.sys >!!win64ddsys.md5
!!md5sum winpmem_mini_x64_rc2.exe >!!winpmem.md5


:: Set Varibles from md5 text file (this includes original file name and formating)
set /p !!arp= < !!arp.md5
set /p !!at= < !!at.md5
set /p !!crypthunter= < !!crypthunter.md5
set /p !!EDD= < !!EDD.md5
set /p !!FDPro= < !!FDPro.md5
::set /p !!getip= < !!getip.md5
set /p !!curl= < !!curl.md5
set /p !!getmac= < !!getmac.md5
set /p !!ipconfig= < !!ipconfig.md5
set /p !!md5sum= < !!md5sum.md5
set /p !!net= < !!net.md5
set /p !!netstat= < !!netstat.md5
set /p !!nbtstat= < !!nbtstat.md5
set /p !!now= < !!now.md5
set /p !!openfiles= < !!openfiles.md5
set /p !!openports= < !!openports.md5
set /p !!psfile= < !!psfile.md5
set /p !!Psinfo= < !!Psinfo.md5
set /p !!pslist= < !!pslist.md5
set /p !!psloggedon= < !!psloggedon.md5
set /p !!route= < !!route.md5
set /p !!Sniffer= < !!Sniffer.md5
set /p !!strings= < !!strings.md5
set /p !!systeminfo= < !!systeminfo.md5
set /p !!WHOAMI= < !!WHOAMI.md5
set /p !!winen= < !!winen.md5
set /p !!winen_= < !!winen_.md5
set /p !!winen64= < !!winen64.md5
set /p !!winen64_= < !!winen64_.md5
::set /p !!driveid= < !!driveid.md5
set /p !!psexec= < !!psexec.md5
set /p !!win32dd= < !!win32dd.md5
set /p !!win32ddsys= < !!win32ddsys.md5
set /p !!win64dd= < !!win64dd.md5
set /p !!win64ddsys= < !!win64ddsys.md5
set /p !!winpmem= < !!winpmem.md5
:: Delete md5 files 
del *.md5

:: Removes first / and other formatting from varible, leaving 32bit hash value.
set !!arp=%!!arp:~1,32%
set !!at=%!!at:~1,32%
set !!crypthunter=%!!crypthunter:~1,32%
set !!EDD=%!!EDD:~1,32%
set !!FDPro=%!!FDPro:~1,32%
::set !!getip=%!!getip:~1,32%
set !!curl=%!!curl:~1,32%
set !!getmac=%!!getmac:~1,32%
set !!ipconfig=%!!ipconfig:~1,32%
set !!md5sum=%!!md5sum:~1,32%
set !!net=%!!net:~1,32%
set !!netstat=%!!netstat:~1,32%
set !!nbtstat=%!!nbtstat:~1,32%
set !!now=%!!now:~1,32%
set !!openfiles=%!!openfiles:~1,32%
set !!openports=%!!openports:~1,32%
set !!psfile=%!!psfile:~1,32%
set !!Psinfo=%!!Psinfo:~1,32%
set !!pslist=%!!pslist:~1,32%
set !!psloggedon=%!!psloggedon:~1,32%
set !!route=%!!route:~1,32%
set !!Sniffer=%!!Sniffer:~1,32%
set !!strings=%!!strings:~1,32%
set !!systeminfo=%!!systeminfo:~1,32%
set !!WHOAMI=%!!WHOAMI:~1,32%
set !!winen=%!!winen:~1,32%
set !!winen_=%!!winen_:~1,32%
set !!winen64=%!!winen64:~1,32%
set !!winen64_=%!!winen64_:~1,32%
::set !!driveid=%!!driveid:~1,32%
set !!psexec=%!!psexec:~1,32%
set !!win32dd=%!!win32dd:~1,32%
set !!win32ddsys=%!!win32ddsys:~1,32%
set !!win64dd=%!!win64dd:~1,32%
set !!win64ddsys=%!!win64ddsys:~1,32%
set !!winpmem=%!!winpmem:~1,32%

::cls
echo !!arp.exe : %!!arp%
echo !!at.exe : %!!at%
echo !!crypthunter.exe : %!!crypthunter%
echo !!EDD.exe : %!!EDD%
echo !!FDPro.exe : %!!FDPro%
::echo !!getip.exe : %!!getip%
echo !!curl.exe : %!!curl%
echo !!getmac.exe : %!!getmac%
echo !!ipconfig.exe : %!!ipconfig%
echo !!md5sum.exe : %!!md5sum%
echo !!net.exe : %!!net%
echo !!netstat.exe : %!!netstat%
echo !!nbtstat.exe : %!!nbtstat%
echo !!now.exe : %!!now%
echo !!openfiles.exe : %!!openfiles%
echo !!openports.exe : %!!openports%
echo !!psfile.exe : %!!psfile%
echo !!Psinfo.exe : %!!Psinfo%
echo !!pslist.exe : %!!pslist%
echo !!psloggedon.exe : %!!psloggedon%
echo !!route.exe : %!!route%
echo !!Sniffer.exe : %!!Sniffer%
echo !!strings.exe : %!!strings%
echo !!systeminfo.exe : %!!systeminfo%
echo !!WHOAMI.exe : %!!WHOAMI%
echo !!winen.exe : %!!winen%
echo winen_.sys : %!!winen_%
echo !!winen64.exe : %!!winen64%
echo winen64_.sys : %!!winen64_%
::echo !!driveid.exe : %!!driveid%
echo !!psexec.exe : %!!psexec%
echo !!win32dd.exe : %!!win32dd%
echo win32dd.sys : %!!win32ddsys%
echo !!win64dd.exe : %!!win64dd%
echo win64dd.sys : %!!win64ddsys%
echo winpmem_mini_x64_rc2.exe : %!!winpmem%
echo.
echo Cross referencing against valid hash values...
echo.

:ARP
find /c /i "%!!arp%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto arpfound
color 0c
echo !!arp.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto AT
:arpfound
echo !!arp.exe verified.

:AT
find /c /i "%!!at%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto atfound
color 0c
echo !!at.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto CRYPTHUNTER
:atfound
echo !!at.exe verified.

:CRYPTHUNTER
find /c /i "%!!crypthunter%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto cryptfound
color 0c
echo !!crypthunter.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
::DELETE following line and remove :: comments if DRIVEID enabled
goto edd
::goto driveid
:cryptfound
echo !!crypthunter.exe verified.

:::driveid
::find /c /i "%!!driveid%" !toolhash.tea > NUL
::if %ERRORLEVEL% == 0 goto driveidfound
::echo !!driveid.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
::goto edd
:::driveidfound
::echo !!driveid.exe verified.

:edd
find /c /i "%!!edd%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto eddfound
echo !!EDDv301.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto fdpro
:eddfound
echo !!EDDv301.exe verified.

:fdpro
find /c /i "%!!fdpro%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto fdprofound
::color 0c
echo !!FDPro.exe HASH MISMATCH OR FILE NOT FOUND. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
::goto getip
goto curl
:fdprofound
echo !!FDPro.exe verified.

:::getip
::find /c /i "%!!getip%" !toolhash.tea > NUL
::if %ERRORLEVEL% == 0 goto getipfound
::color 0c
::echo !!getip.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
::goto getmac
:::getipfound
::echo !!getip.exe verified.

:curl
find /c /i "%!!curl%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto curlfound
color 0c
echo !!curl.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto getmac
:curlfound
echo !!curl.exe verified.

:getmac
find /c /i "%!!getmac%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto getmacfound
color 0c
echo !!getmac.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto ipconfig
:getmacfound
echo !!getmac.exe verified.

:ipconfig
find /c /i "%!!ipconfig%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto ipconfigfound
color 0c
echo !!ipconfig.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto md5sum
:ipconfigfound
echo !!ipconfig.exe verified.

:md5sum
find /c /i "%!!md5sum%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto md5sumfound
color 0c
echo !!md5sum.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto net
:md5sumfound
echo !!md5sum.exe verified.

:net
find /c /i "%!!net%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto netfound
color 0c
echo !!net.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto netstat
:netfound
echo !!net.exe verified.

:netstat
find /c /i "%!!netstat%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto netstatfound
color 0c
echo !!netstat.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto nbtstat
:netstatfound
echo !!netstat.exe verified.

:nbtstat
find /c /i "%!!nbtstat%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto nbtstatfound
color 0c
echo !!nbtstat.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto now
:nbtstatfound
echo !!nbtstat.exe verified.

:now
find /c /i "%!!now%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto nowfound
color 0c
echo !!now.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto openfiles
:nowfound
echo !!now.exe verified.

:openfiles
find /c /i "%!!openfiles%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto openfilesfound
color 0c
echo !!openfiles.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto openports
:openfilesfound
echo !!openfiles.exe verified.

:openports
find /c /i "%!!openports%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto openportsfound
color 0c
echo !!openports.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto psfile
:openportsfound
echo !!openports.exe verified.

:psfile
find /c /i "%!!psfile%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto psfilefound
color 0c
echo !!psfile.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto Psinfo
:psfilefound
echo !!psfile.exe verified.


:psinfo
find /c /i "%!!psinfo%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto psinfofound
echo !!psinfo.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
color 0c
goto pslist
:psinfofound
echo !!psinfo.exe verified.

:pslist
find /c /i "%!!pslist%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto pslistfound
color 0c
echo !!pslist.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto psloggedon
:pslistfound
echo !!pslist.exe verified.

:psloggedon
find /c /i "%!!psloggedon%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto psloggedonfound
color 0c
echo !!psloggedon.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto route
:psloggedonfound
echo !!psloggedon.exe verified.

:route
find /c /i "%!!route%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto routefound
color 0c
echo !!route.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto sniffer
:routefound
echo !!route.exe verified.

:sniffer
find /c /i "%!!sniffer%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto snifferfound
color 0c
echo !!sniffer.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto strings
:snifferfound
echo !!sniffer.exe verified.

:strings
find /c /i "%!!strings%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto stringsfound
color 0c
echo !!strings.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto systeminfo
:stringsfound
echo !!strings.exe verified.

:systeminfo
find /c /i "%!!systeminfo%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto systeminfofound
color 0c
echo !!systeminfo.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto whoami
:systeminfofound
echo !!systeminfo.exe verified.

:whoami
find /c /i "%!!whoami%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto whoamifound
color 0c
echo !!whoami.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto winen
:whoamifound
echo !!whoami.exe verified.

:winen
find /c /i "%!!winen%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto winenfound
color 0c
echo !!winen.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto winensys
:winenfound
echo !!winen.exe verified.

:winensys
find /c /i "%!!winen_%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto winensysfound
color 0c
echo winen_.sys HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto winen64
:winensysfound
echo winen_.sys verified.

:winen64
find /c /i "%!!winen64%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto winen64found
color 0c
echo !!winen64.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto winen64sys
:winen64found
echo !!winen64.exe verified.

:winen64sys
find /c /i "%!!winen64_%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto winen64sysfound
color 0c
echo winen64_.sys HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto psexec
:winen64sysfound
echo winen64_.sys verified.

:psexec
find /c /i "%!!psexec%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto psexecfound
color 0c
echo !!psexec.exe HASH MISMATCH. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto win32dd
:psexecfound
echo !!psexec.exe verified.


:win32dd
find /c /i "%!!win32dd%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto win32ddfound
::color 0c
echo !!win32dd.exe HASH MISMATCH OR FILE NOT FOUND. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto !!win32ddsys
:win32ddfound
echo !!win32dd.exe verified.

:!!win32ddsys
find /c /i "%!!win32ddsys%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto !!win32ddsysfound
::color 0c
echo win32dd.sys HASH MISMATCH OR FILE NOT FOUND. PLEASE UPDATE TOOL OR !TOOLHASH.TEA  
goto !!win64dd
:!!win32ddsysfound
echo win32dd.sys verified.

:win64dd
find /c /i "%!!win64dd%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto win64ddfound
::color 0c
echo !!win64dd.exe HASH MISMATCH OR FILE NOT FOUND. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto !!win64ddsys
:win64ddfound
echo !!win64dd.exe verified.

:!!win64ddsys
find /c /i "%!!win64ddsys%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto !!win64ddsysfound
::color 0c
echo win64dd.sys HASH MISMATCH OR FILE NOT FOUND. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto winpmem
:!!win64ddsysfound
echo win64dd.sys verified.

:winpmem
find /c /i "%!!winpmem%" !toolhash.tea > NUL
if %ERRORLEVEL% == 0 goto winpmemfound
::color 0c
echo winpmem_mini_x64_rc2.exe HASH MISMATCH OR FILE NOT FOUND. PLEASE UPDATE TOOL OR !TOOLHASH.TEA 
goto exit
:winpmemfound
echo winpmem_mini_x64_rc2.exe verified.


:exit
echo.
echo ### GREEN text shows all tools verified. ###
echo !!! RED text shows some tools hash dont match or missing!!!
echo.
echo.
echo Colours are not applied to DriveID, FDPro or Moonsols as they are not 
echo included with TEA distribution.
echo.
pause
color 07