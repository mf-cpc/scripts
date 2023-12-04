@echo off
:: ######################################################
:: VERSION 0.9.121221a - December 21 2012 Apocolypse Edition
:: SCRIPT: Batch Run Encapsulation Widget (BREW)
:: CREATION DATE: 2011-02-28
:: LAST MODIFIED: 2012-12-21
:: AUTHOR: Cst. Mark SOUTHBY, #53384
:: EMAIL: mark.southby@rcmp-grc.gc.ca
:: ######################################################
:: #### Authored by Cst. SOUTHBY, #53384 of "E" Div. Pacific Region Integrated Tech Crimes Unit - Island District. 
:: ---------------------------------------------- 
:: #### For Law enforcement use only. Can be used and modified without permission for Law Enforcement purposes.
:: DESCRIPTION:TEA configuration generator and tool checker
:: START OF BATCH
:begin
:: Set color (Green text)
cls
title BREWING TEA
color 0e
echo ==-Batch Run Encapsulation Widget for Tech Evidence Aquisition (BREW TEA)-==
:check
IF EXIST tea.cfg goto found
set runip=0
set rundns=0
set runarp=0
set runnbt=0
set runroute=0
set runnetstat=0
set runopenports=0
set runnetview=0
set runnetshare=0
set runnetses=0
set runpsloggedon=0
set runpslist=0
set runpsfile=0
set runsysinfo=0
set runpsinfo=0
set runwhoami=0
set runat=0
set rungetmac=0
set rungetip=0
set runcurl=0
set rundriveid=0
set cffound=0
set delcfg=1
:: ---- new varibles
set filenum=0
set exnum=0
set invest=NONE
set mem=0
set cap=0
set windd=0
set winencap=0
set fdp=0


:: ===================================================================================
:: if %runip%==y echo runip=1 >>tea.cfg
:: ===================================================================================
:cls
echo Checking for memory capture tools...
:moonsols
:: UPDATE FOR VER 0.9.111124 FDPRO replaced with Moonsols WinDD
::Check if !!win32dd is on USB. if so, set varible to one, else go to message
if exist !!win32dd.exe (set windd=1) Else (goto NOWINDD)
set cap=1
goto setfile

:: UPDATE FOR VER 0.9.111124 FDPRO replaced with Moonsols WinDD
:NOWINDD
:: echo Moonsols WinDD not found. Check for FDPro? (Skip to use WinEn memory capture).
:: set /p winencap= Check for FDPro? (Y)es or (N)o :
:: disables error message when nothing selected
set mem=2
set cap=0
color 4e
echo.
echo A copy of Moonsols was not found!
echo.
echo To capture memory and the page file using TEA
echo You must have a licensed copy of
echo Moonsols or FDPro.
echo.
echo Once you have purchased the software
echo copy win32dd.exe, win32dd.sys,win32dd.exe, win32dd.sys
echo to the same folder as this batch file (TEA)
echo rename the exe files to have !! as a prefix (!!win32dd.exe)
echo.
echo To cancel now, press CTRL-C,
echo.
echo.
set /p winencap= Would you like to use (F)DPro or (W)inEnCap :
if %winencap%==w set cap=3
if %winencap%==W set cap=3
if %winencap%==f goto fdpro
if %winencap%==F goto fdpro
goto setfile


:fdpro
::Check if !!FDPro is on USB. if so, set varible to one, else go to message
if exist !!FDPro.exe (set fdp=1) Else (goto NOFDPRO)
set cap=2
goto setfile

:NOFDPRO
:: Set fdp varible to 0, this will allow WINEN to run
set fdp=0
:: disables error message when nothing selected
set mem=2
color 4e
echo.
echo A copy of FDPro was not found!
echo.
echo To capture memory and the page file using TEA
echo You must have a licensed copy of
echo Moonsols or FDPro.
echo.
echo Once you have purchased the software
echo copy win32dd.exe, win32dd.sys,win32dd.exe, win32dd.sys
echo to the same folder as this batch file (TEA)
echo rename the exe files to have !! as a prefix (!!win32dd.exe)
echo.
echo To cancel now, press CTRL-C,
echo.
echo Would you like to proceed using WINEN to perform a Memory Capture only
echo.
set /p winencap= (Y)es or (N)o:
if %winencap%==y set cap=3
if %winencap%==Y set cap=3
if %winencap%==n set cap=4
if %winencap%==N set cap=4
goto setfile


:: -----------------------------------------------------
:setfile
cls
echo.
echo	 =============================================
echo	 #                                           #
echo	 # PACIFIC REGION INTEGRATED TECH CRIME UNIT #
echo	 #                                           #
echo	 #      Batch Run Encapsulation Widget for   #
echo	 #      Technological Evidence Aquisition    #
echo	 #                                           #
echo	 #            (B.R.E.W. T.E.A.)              #
echo	 #                                           #
echo	 =============================================
echo.
echo.
echo	    NO SPACES OR SPECIAL CHARACTERS EXCEPT   
echo	                    - _ . !                  
echo.

:: Ask for file, exhibit and investigator. If left blank, go to varible "blank" which is error message.
set /p filenum=File Number: 
if %filenum% == 0 goto blank
echo filenum=%filenum%>>tea.cfg

set /p exnum=Exhibit # or 1 Word Description: 
if %exnum% == 0 goto blank
echo exnum=%exnum%>>tea.cfg
IF EXIST %filenum%-%exnum%/ goto exst
set /p invest=Investigator Name (Spaces Allowed): 
echo invest=%invest% >>tea.cfg
:: Supposed to prevent nothing being entered into investigator name, however produced error if a space was present. 
::if %invest% == 0 goto blank
echo.
echo Select Memory Capture Options:
if %cap% == 1 set /p mem=(R)aw memory only, with (B)IOS, All w/(D)evice memory [MAY CRASH], (N)one:
if %cap% == 2 set /p mem=FDPRO:(M)em Only; (P)age file and mem; (N)one:
if %cap% == 3 set /p mem=WinEnCap:(C)apture memory; (N)one:
if %mem% == 0 goto blank
echo mem=%mem% >>tea.cfg
goto tools

:blank
color e4
cls
echo FIELD CAN NOT BE LEFT BLANK!
pause
goto setfile

:exst
color 0c
cls
echo. 
echo !!! Folder %filenum%-%exnum% exists! 
echo Please type a different exhibit number or delete old folder!
pause
goto setfile
:: -----------------------------------------------------
:: ===================================================================================
:: ===================================================================================
:tools
echo Set up Tools:
echo.
echo.
echo Please select the applications you wish to run
echo * indicates will run by default if no config file present.
echo Type (Y) or (N) to configure TEA:
goto rip

:rip
set /p runip=IPCONFIG /ALL [Network Settings] *:
if %runip%==y echo runip=1 >>tea.cfg
if %runip%==Y echo runip=1 >>tea.cfg
goto rdns

:rdns
set /p rundns=IPCONFIG /DISPLAYDNS [Show Which Sites Visited] *:
if %rundns%==y echo rundns=1 >>tea.cfg
if %rundns%==Y echo rundns=1 >>tea.cfg
goto rarp

:rarp
set /p runarp=ARP -A [MAC Connections]:
if %runarp%==y echo runarp=1 >>tea.cfg
if %runarp%==Y echo runarp=1 >>tea.cfg
goto rnbt

:rnbt
set /p runnbt=NBTSTAT -N [Netbios Names in cache]:
if %runnbt%==y echo runnbt=1 >>tea.cfg
if %runnbt%==Y echo runnbt=1 >>tea.cfg
goto rroute

:rroute
set /p runroute=ROUTE PRINT [Show Routing Tables]:
if %runroute%==y echo runroute=1 >>tea.cfg
if %runroute%==Y echo runroute=1 >>tea.cfg
goto rnetstat

:rnetstat
set /p runnetstat=NETSTAT -ABON [ip and ports belonging to proccess']*:
if %runnetstat%==y echo runnetstat=1 >>tea.cfg
if %runnetstat%==Y echo runnetstat=1 >>tea.cfg
goto ropenports

:ropenports
set /p runopenports=OPENPORTS -LINES -PATH [similar to netstat, shows full path on XP]:
if %runopenports%==y echo runopenports=1 >>tea.cfg
if %runopenports%==Y echo runopenports=1 >>tea.cfg
goto rnetview

:rnetview
set /p runnetview=NET VIEW [Discoverable devices on network] *:
if %runnetview%==y echo runnetview=1 >>tea.cfg
if %runnetview%==Y echo runnetview=1 >>tea.cfg
goto rnetshare

:rnetshare
set /p runnetshare=NET SHARE [Open Shares] *:
if %runnetshare%==y echo runnetshare=1 >>tea.cfg
if %runnetshare%==Y echo runnetshare=1 >>tea.cfg
goto rnetses

:rnetses
set /p runnetses=NET SESSIONS [Open network sessions]:
if %runnetses%==y echo runnetses=1 >>tea.cfg
if %runnetses%==Y echo runnetses=1 >>tea.cfg
goto rpsloggedon

:rpsloggedon
set /p runpsloggedon=PSLOGGEDON /ACCEPTEULA [Whos logged on]:
if %runpsloggedon%==y echo runpsloggedon=1 >>tea.cfg
if %runpsloggedon%==Y echo runpsloggedon=1 >>tea.cfg
goto rpslist

:rpslist
set /p runpslist=PSLIST -T [Processes running]:
if %runpslist%==y echo runpslist=1 >>tea.cfg
if %runpslist%==Y echo runpslist=1 >>tea.cfg
goto rpsfile

:rpsfile
set /p runpsfile=PSFILE /ACCEPTEULA [Shows files opened remotely]:
if %runpsfile%==y echo runpsfile=1 >>tea.cfg
if %runpsfile%==Y echo runpsfile=1 >>tea.cfg
goto rsysinfo

:rsysinfo
set /p runsysinfo=SYSTEMINFO [System Info] *:
if %runsysinfo%==y echo runsysinfo=1 >>tea.cfg
if %runsysinfo%==Y echo runsysinfo=1 >>tea.cfg
goto rpsinfo

:rpsinfo
set /p runpsinfo=PSINFO -H /ACCEPTEULA [Shows system patches. sysinfo shows this too]:
if %runpsinfo%==y echo runpsinfo=1 >>tea.cfg
if %runpsinfo%==Y echo runpsinfo=1 >>tea.cfg
goto rwhoami

:rwhoami
set /p runwhoami=WHOAMI /USER /SID [user account name and SID]:
if %runwhoami%==y echo runwhoami=1 >>tea.cfg
if %runwhoami%==Y echo runwhoami=1 >>tea.cfg
goto rat

:rat
set /p runat=AT [Scheduled tasks]:
if %runat%==y echo runat=1 >>tea.cfg
if %runat%==Y echo runat=1 >>tea.cfg
goto rgetmac

:rgetmac
set /p rungetmac=GETMAC /V [Shows MAC address and transport name]:
if %rungetmac%==y echo rungetmac=1 >>tea.cfg
if %rungetmac%==Y echo rungetmac=1 >>tea.cfg
goto rcurl

:rcurl
set /p runcurl=CURL [Shows what the external IP address is shown as] *:
if %runcurl%==y echo runcurl=1 >>tea.cfg
if %runcurl%==Y echo runcurl=1 >>tea.cfg
goto rdriveid

:rdriveid
set /p rundriveid=EDD (or DriveID) [Shows Encryption] *:
if %rundriveid%==y echo rundriveid=1 >>tea.cfg
if %rundriveid%==Y echo rundriveid=1 >>tea.cfg


IF EXIST tea.cfg goto found
echo ==- TEA CONFIGURATION NOT GENERATED -==
goto check


:found
set cffound=1
:dlist
FOR /F "tokens=2 delims==" %%a IN ('find "runip" ^<tea.cfg') DO SET runip=%%a
FOR /F "tokens=2 delims==" %%t IN ('find "rundns" ^<tea.cfg') DO SET rundns=%%t
FOR /F "tokens=2 delims==" %%b IN ('find "runarp" ^<tea.cfg') DO SET runarp=%%b
FOR /F "tokens=2 delims==" %%c IN ('find "runnbt" ^<tea.cfg') DO SET runnbt=%%c
FOR /F "tokens=2 delims==" %%d IN ('find "runroute" ^<tea.cfg') DO SET runroute=%%d
FOR /F "tokens=2 delims==" %%e IN ('find "runnetstat" ^<tea.cfg') DO SET runnetstat=%%e
FOR /F "tokens=2 delims==" %%f IN ('find "runopenports" ^<tea.cfg') DO SET runopenports=%%f
FOR /F "tokens=2 delims==" %%g IN ('find "runnetview" ^<tea.cfg') DO SET runnetview=%%g
FOR /F "tokens=2 delims==" %%h IN ('find "runnetshare" ^<tea.cfg') DO SET runnetshare=%%h
FOR /F "tokens=2 delims==" %%i IN ('find "runnetses" ^<tea.cfg') DO SET runnetses=%%i
FOR /F "tokens=2 delims==" %%j IN ('find "runpsloggedon" ^<tea.cfg') DO SET runpsloggedon=%%j
FOR /F "tokens=2 delims==" %%k IN ('find "runpslist" ^<tea.cfg') DO SET runpslist=%%k
FOR /F "tokens=2 delims==" %%l IN ('find "runpsfile" ^<tea.cfg') DO SET runpsfile=%%l
FOR /F "tokens=2 delims==" %%m IN ('find "runsysinfo" ^<tea.cfg') DO SET runsysinfo=%%m
FOR /F "tokens=2 delims==" %%n IN ('find "runpsinfo" ^<tea.cfg') DO SET runpsinfo=%%n
FOR /F "tokens=2 delims==" %%o IN ('find "runwhoami" ^<tea.cfg') DO SET runwhoami=%%o
FOR /F "tokens=2 delims==" %%p IN ('find "runat" ^<tea.cfg') DO SET runat=%%p
FOR /F "tokens=2 delims==" %%q IN ('find "rungetmac" ^<tea.cfg') DO SET rungetmac=%%q
::FOR /F "tokens=2 delims==" %%r IN ('find "rungetip" ^<tea.cfg') DO SET rungetip=%%r
FOR /F "tokens=2 delims==" %%s IN ('find "rundriveid" ^<tea.cfg') DO SET rundriveid=%%s
FOR /F "tokens=2 delims==" %%u IN ('find "runcurl" ^<tea.cfg') DO SET runcurl=%%u
:: ----
FOR /F "tokens=2 delims==" %%1 IN ('find "filenum" ^<tea.cfg') DO SET filenum=%%1
FOR /F "tokens=2 delims==" %%2 IN ('find "exnum" ^<tea.cfg') DO SET exnum=%%2
FOR /F "tokens=2 delims==" %%3 IN ('find "invest" ^<tea.cfg') DO SET invest=%%3
FOR /F "tokens=2 delims==" %%4 IN ('find "mem" ^<tea.cfg') DO SET mem=%%4
FOR /F "tokens=2 delims==" %%5 IN ('find "cap" ^<tea.cfg') DO SET cap=%%5
FOR /F "tokens=2 delims==" %%6 IN ('find "windd" ^<tea.cfg') DO SET windd=%%6
FOR /F "tokens=2 delims==" %%7 IN ('find "winencap" ^<tea.cfg') DO SET winencap=%%7
FOR /F "tokens=2 delims==" %%8 IN ('find "fdp" ^<tea.cfg') DO SET fdp=%%8
::goto delcfg
:redisplay
cls
echo ==- TEA CONFIGURATION FILE EXISTS... -==
echo.
echo = CURRENT TEA CONFIGURATION =
echo.
echo File Number:		%filenum%
echo Exhibit Number:		%exnum%
echo Investigator:		%invest%
if %mem% == m echo FDPRO MEMORY CAPTURE: Memory dump Only
if %mem% == M echo FDPRO MEMORY CAPTURE: Memory dump Only
if %mem% == p echo FDPRO MEMORY CAPTURE: Memory and Pagefile dump
if %mem% == P echo FDPRO MEMORY CAPTURE: Memory and Pagefile dump
if %mem% == r echo MOONSOLS MEMORY CAPTURE: Memory dump Only
if %mem% == R echo MOONSOLS MEMORY CAPTURE: Memory dump Only
if %mem% == b echo MOONSOLS MEMORY CAPTURE: RAW and BIOS memory dump
if %mem% == B echo MOONSOLS MEMORY CAPTURE: RAW and BIOS memory dump
if %mem% == D echo MOONSOLS MEMORY CAPTURE: All Memory [May cause BSOD CRASH!]
if %mem% == d echo MOONSOLS MEMORY CAPTURE: All Memory [May cause BSOD CRASH!]
if %mem% == c echo WINENCAP MEMORY CAPTURE: Memory dump Only
if %mem% == C echo WINENCAP MEMORY CAPTURE: Memory dump Only
if %mem% == N echo MEMORY CAPTURE: None
if %mem% == n echo MEMORY CAPTURE: None
if %mem% == N echo MEMORY CAPTURE: None
echo ------ Tools ------
::if %runpwdump%==1 echo -	PWDUMP7				[Dumps User account password hashes]
if %runip%==1 echo -	IPCONFIG /ALL			[Network Settings]
if %rundns%==1 echo -	IPCONFIG /DISPLAYDNS		[Show Which Sites Visited]
if %runarp%==1 echo - 	ARP -A				[MAC Connections]
if %runnbt%==1 echo -	NBTSTAT -N			[Netbios Names in cache]
if %runroute%==1 echo -	ROUTE PRINT			[Show Routing Tables]
if %runnetstat%==1 echo -	NETSTAT -ABON			[ip and ports belonging to proccess']
if %runopenports%==1 echo -	OPENPORTS -LINES -PATH		[similar to netstat, shows full path on XP]
if %runnetview%==1 echo -	NET VIEW			[Discoverable devices on network]
if %runnetshare%==1 echo -	NET SHARE			[Open Shares]
if %runnetses%==1 echo -	NET SESSIONS			[Open network sessions]
if %delcfg%==r pause
if %delcfg%==R pause
if %runpsloggedon%==1 echo -	PSLOGGEDON /ACCEPTEULA		[Whos logged on]
if %runpslist%==1 echo -	PSLIST -T			[Processes running]
if %runpsfile%==1 echo -	PSFILE /ACCEPTEULA		[Shows files opened remotely]
if %runsysinfo%==1 echo -	SYSTEMINFO			[System Info] 
if %runpsinfo%==1 echo -	PSINFO -H /ACCEPTEULA		[Shows system patches. sysinfo shows this too]
if %runwhoami%==1 echo -	WHOAMI /USER /SID		[user account name and SID]
if %runat%==1 echo -	AT				[Scheduled tasks]
if %rungetmac%==1 echo -	GETMAC /V			[Shows MAC address and transport name]
::if %rungetip%==1 echo -	GETIP				[Shows what the external IP address is shown as]
if %runcurl%==1 echo -	CURL				[Shows what the external IP address is shown as] 			
if %rundriveid%==1 echo -	EDD (or DriveID)		[Shows Encryption]
if %cffound%==1 goto delcfg
pause
:: goto exit
:delcfg
echo.
echo.
echo -= You can(R)edisplay configuration with a pause half way. =-
echo.
echo Would you like to keep the current configuration?
set /p delcfg=(Y)es or (N)o or (R)edisplay :
if %delcfg%==n goto dcfg
if %delcfg%==N goto dcfg
if %delcfg%==r goto redisplay
if %delcfg%==R goto redisplay
goto exit

:dcfg
del tea.cfg
set cffound=0
goto begin

:exit
pause
verify_tools.cmd