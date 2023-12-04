@echo off
:: ######################################################
:: VERSION 0.9.130107a - January 2013 HoNeY Edition (HNY=Happy New Year)
set ver=0.9.130107a
:: SCRIPT: Technological Evidence Acquisition (TEA)
:: CREATION DATE: 2009-12-04
:: LAST MODIFIED: 2012-01-07
:: Hash list current as of: 
:: AUTHOR: Cst. Mark SOUTHBY
:: EMAIL: mark@southby.ca
:: ######################################################
:: DESCRIPTION: A batch file created to run numerous tools
:: useful for collecting information on a running (live) system.
:: This tool will run all tools from the USB (XP Mode) if 
:: the OS version is detected as: 95, 98, ME, XP, NT, 2000, 2003.
:: Windows Vista/7 compatibility: The newer versions of windows
:: prevent system tools being run outside of the Windows/System32 
:: directory. This tool will hash those tools and compare them to
:: a list of known hash values for those applications. 
:: This hash list is a text file called "!hashlist.tea".
:: If the OS version isn't detected as XP Mode and the Hash values 
:: do not match, then those tools will be skipped. 
:: You can always run them manually at the scene.
:: PLEASE NOTE: Works with FDPro to capture memory and page files (or just memory)
:: FDPro is licensed software from HBGary. You must purchase a copy of FDPro first. 
:: Copy FDPro.exe to same folder as this batch file and rename to !!FDPro.exe
:: memory and a page file. Use !!EDD to image a drive/device.
:: ######################################################
:: KNOWN ISSUES:
:: (1) Currently do not have updated Hashes for Windows 7 SP1 tools nor Windows 8 tools
:: (2) The version of DRIVEID I have does not detect bitlocker, so it has been renamed to !driveid.exe . 
::     If you wish to still run DRIVEID, rename it to !!driveid.exe or put a more current version in this folder.
::     Dont forget to add the MD5 hash of the next executable to !toolhash.tea
:: (3) It is rather difficult to add new tools or change the order of tools, if you really need it, 
::     email me and I will do my best to update it. I am hoping to expand my programming skills and reprogram 
::     this tool to include this option.
::     as this version was originally created for my own personal use it isn't polished.
:: (4) Most of these tools have been tested in both Windows 7 and XP environments. 
::     If you find issues on these or other OS's, please let me know. 
::     If you successfully use this tool, feedback would be appreciated.
:: (5) In tests, PWDUMP tools have issues with virus scanners and UAC on XP sp3. Disabled till further tests conducted.
:: ######################################################
:: #### Authored by Cst. SOUTHBY, for personal use. Not sanctioned or approved by any organization or entity. Use at your own risk.
:: #### For Law enforcement use only. Please read "README_TEA.RTF" for End User Licence Agreement 
:: You can edit it and give it to law enforcement. Please give credit where credit is due.
:: Don't disclose or give copies to non-tech crime persons, as it is in the same realm of other forensic tools and 
:: considered investigational techniques. Reports generated are disclosable. 

::Delcare and clear all varibles used
set archi=0
set !!user=0
set admin=0
set filenum=0
set exnum=0
set invest=0
set mem=0
set fdp=0
set arch=0
set ipconfig=0
set ipc=0
set at=0
set route=0
set netstat=0
set net=0
set arp=0
set cap=0
set cnet=0
set runpwdump=0
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
set runwin32dd=0
set runwin64dd=0
set continue=0

cls
:: Figuring out if Windows XP / 7, 32bit or 64bit!
if %PROCESSOR_ARCHITECTURE%==x86 set archi=32
if %archi%==32 goto admin
set archi=64
goto admin

:admin
:: START OF ADMIN CHECK
TITLE Admin Check
echo ======================================================================= 
echo      Check to see if User accout is part of Administator Group
echo ======================================================================= 
color 0c
!!whoami > usr.tmp
set /p !!user= < usr.tmp
echo.
net user %username% | findstr /r Administrator
echo.
if %errorlevel% == 1 (goto checksys) else (goto admin)
:checksys
echo.
type usr.tmp|findstr /r AUTHORITY
if %errorlevel% == 1 (goto notpriv) else (goto system)
goto system

:notpriv
echo Currently running command prompt as: %!!user%
echo.
echo.
echo **WARNING** This is does not appear to be part of the local Administators or SYSTEM group... 
echo.
echo.
echo.
echo.
if %archi%==64 set /p admin= Does the title of this window say "Administrator: Admin Check" (Y)es or (N)o:
if %admin%==y goto start
if %admin%==Y goto start
echo Consider running "escalate.cmd"
echo Press Ctrl-C now to cancel, or 
pause
del usr.tmp
goto admininstr

:system
color 0a
echo.
echo Running TEA under: %!!user%
echo.
echo  ======================================================================
echo  # You are running as SYSTEM and should not have any permision issues.#
echo  ======================================================================
echo.
pause
goto start

:admin
color 0a
echo.
echo Running TEA under: %!!user%
echo.
echo        =============================
echo        # This IS an admin account! #
echo        =============================
echo.
echo.
if %archi%==64 set /p admin= Does the title of this window say "Administrator: Admin Check" (Y)es or (N)o:
if %admin%==y goto start
if %admin%==Y goto start
if %archi%==32 goto start
goto admininstr

:start
IF EXIST usr.tmp del usr.tmp
:: clear screen
cls
::title
TITLE Technological Evidence Aquisition
color 0a
cls
echo.
echo.
echo.	                                            
echo	        Technological Evidence Aquisition    
echo	                   (T.E.A.)                  
echo.
echo               Checking Configuration...


::Check for Config
:cnfchk
IF EXIST tea.cfg goto found
goto notfound

:found
FOR /F "tokens=2 delims==" %%a IN ('find "runip" ^<tea.cfg') DO SET runip=%%a
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
FOR /F "tokens=2 delims==" %%r IN ('find "rungetip" ^<tea.cfg') DO SET rungetip=%%r
FOR /F "tokens=2 delims==" %%s IN ('find "rundriveid" ^<tea.cfg') DO SET rundriveid=%%s
FOR /F "tokens=2 delims==" %%t IN ('find "rundns" ^<tea.cfg') DO SET rundns=%%t
FOR /F "tokens=2 delims==" %%u IN ('find "runcurl" ^<tea.cfg') DO SET runcurl=%%u
FOR /F "tokens=2 delims==" %%v IN ('find "runpwdump" ^<tea.cfg') DO SET runpwdump=%%v
FOR /F "tokens=2 delims==" %%w IN ('find "runwin32dd" ^<tea.cfg') DO SET runwin32dd=%%w
FOR /F "tokens=2 delims==" %%x IN ('find "runwin64dd" ^<tea.cfg') DO SET runwin64dd=%%x
goto begin

:notfound
Echo No config file found. Would you like to run default tools?
echo [IPCONFIG, NETSTAT, NET SHARE, NET VIEW, SYSTEMINFO, GET EXTERNAL IP, CHECK for ENCRYPTION]
set /p rundefault=(Y)es or (N)o:
if %rundefault%==y goto rndefault
if %rundefault%==Y goto rndefault
goto brew

:rndefault
set runpwdump=1
set runip=1
set rundns=1
set runarp=0 
set runnbt=0 
set runroute=0 
set runnetstat=1 
set runopenports=0 
set runnetview=1 
set runnetshare=1 
set runnetses=0
set runpsloggedon=0 
set runpslist=0 
set runpsfile=0 
set runsysinfo=1 
set runpsinfo=0 
set runwhoami=0 
set runat=0 
set rungetmac=0 
::set rungetip=1 
set runcurl=1 
set rundriveid=1 
goto begin

::RUN NEXT
::XP MODE
:run0a
::if %runpwdump%==1 goto pwdmp7
:run00
if %runip%==1 goto xpip
:run00a
if %rundns%==1 goto xpdns
:run01
if %runarp%==1 goto xparp
:run02
if %runroute%==1 goto xproute
:run03
if %runnetstat%==1 goto xpnetstat
:run04
if %runnetses%==1 goto xpnetses
:run05
if %runnetview%==1 goto xpnetview
:run06
if %runnetshare%==1 goto xpnetshare
:run07
if %runat%==1 goto xpat
::ONCE XP TOOLS RUN, GOTO RUN ALL
goto run20

::WIN 7
:run09
::if %runpwdump%==1 goto pwdump7
:cip
if %runip%==1 goto checkip
if %rundns%==1 goto checkip
:run10
if %runip%==1 goto ipfound
:run10a
if %rundns%==1 goto checkdns
:run11
if %runat%==1 goto checkat
:run12
if %runroute%==1 goto checkroute
:run13
if %runnetstat%==1 goto checknetstat
:cnet
if %runnetses%==1 goto checknet
:run14
if %runnetses%==1 goto netfound 	
:run15
if %runnetview%==1 goto 7netview
:run16
if %runnetshare%==1 goto 7netshare
:run17
if %runarp%==1 goto checkarp
::ONCE WIN 7 TOOLS RUN, GOTO RUN ALL
goto run20

::Run ALL TOOLS
:run20
if %runnbt%==1 goto rnnbtstat
:run21
if %runopenports%==1 goto rnopenports
:run22
if %runpsloggedon%==1 goto rnpsloggedon
:run23
if %runpslist%==1 goto rnpslist
:run24
if %runpsfile%==1 goto rnpsfile
:run25
if %runsysinfo%==1 goto rnsysinfo
:run26
if %runpsinfo%==1 goto rnpsinfo
:run27
if %runwhoami%==1 goto rnwhoami
:run28
if %rungetmac%==1 goto rngetmac
:::run29
::if %rungetip%==1 goto rngetip
:run29
if %runcurl%==1 goto rncurl
:run29B
if %rundriveid%==1 goto rndriveid
:run30
goto endfile

:: ### START OF FILE ####
:begin
:: Clear screen
cls

:moonsols
:: UPDATE FOR VER 0.9.111124 FDPRO replaced with Moonsols WinDD
::Check if !!win32dd is on USB. if so, set varible to one, else go to message
if exist !!win32dd.exe (set windd=1) Else (goto NOWINDD)
set cap=1
goto setfile

:: UPDATE FOR VER 0.9.111124 FDPRO replaced with Moonsols WinDD
:NOWINDD
:: Set fpd varible to 0, this will allow WINEN to run
echo Moonsols WinDD not found. Check for FDPro? (Skip for WinEn Cap).
set /p winencap= Check for FDPro? (Y)es or (N)o :
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
:: Set fpd varible to 0, this will allow WINEN to run
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

:exst
color 0c
cls
echo. 
echo !!! Folder %filenum%-%exnum% exists! 
echo Please type a different exhibit number or delete old folder!
pause
goto setfile

:setfile
:: Set color (Green text)
color 0a
cls
echo.
echo	 =============================================
echo	 #                                           #
echo	 #                                           #
echo	 #      Technological Evidence Aquisition    #
echo	 #                 (T.E.A.)                  #
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
set /p exnum=Exhibit # or 1 Word Description: 
if %exnum% == 0 goto blank
IF EXIST %filenum%-%exnum%/ goto exst
set /p invest=Investigator Name (Spaces Allowed): 
:: Supposed to prevent nothing being entered into investigator name, however produced error if a space was present. 
::if %invest% == 0 goto blank
echo.
echo Select Memory Capture Options:
if %cap% == 1 set /p mem=(R)aw memory only, with (B)IOS, or {E}veryting [BIOS, Video, RAM, MAY CRASH], (N)one:
if %cap% == 2 set /p mem=FDPRO:(M)em Only; (P)age file and mem; (N)one:
if %cap% == 3 set /p mem=Dumpit: (D)ump memory; (N)one;
if %cap% == 4 set /p mem=WinEnCap:(C)apture memory; (N)one:
if %mem% == 0 goto blank

:: clear screen
cls

:: confirm data input correct
echo ==- LIVE ACQUISITION -==
echo. 
echo File#:        %filenum%
echo Exhibit#:     %exnum%
echo Investigator: %invest%
echo.
if %mem% == m echo FDPRO MEMORY CAPTURE: Memory dump Only
if %mem% == M echo FDPRO MEMORY CAPTURE: Memory dump Only
if %mem% == p echo FDPRO MEMORY CAPTURE: Memory and Pagefile dump
if %mem% == P echo FDPRO MEMORY CAPTURE: Memory and Pagefile dump
if %mem% == r echo MOONSOLS MEMORY CAPTURE: Memory dump Only
if %mem% == R echo MOONSOLS MEMORY CAPTURE: Memory dump Only
if %mem% == b echo MOONSOLS MEMORY CAPTURE: RAW and BIOS memory dump
if %mem% == B echo MOONSOLS MEMORY CAPTURE: RAW and BIOS memory dump
if %mem% == e echo MOONSOLS MEMORY CAPTURE: All Memory [May cause BSOD CRASH!]
if %mem% == E echo MOONSOLS MEMORY CAPTURE: All Memory [May cause BSOD CRASH!]
if %mem% == d echo MOONSOLS DUMPIT : Memory dump Only
if %mem% == D echo MOONSOLS DUMPIT : Memory dump Only
if %mem% == c echo WINENCAP MEMORY CAPTURE: Memory dump Only
if %mem% == C echo WINENCAP MEMORY CAPTURE: Memory dump Only
if %mem% == n echo MEMORY CAPTURE: None
if %mem% == N echo MEMORY CAPTURE: None
:: If any other letter or number typed in other than M, P or N, it will do the same as N but not display that none was selected.
echo.
set /p continue=Is the above correct? (Y)es to start or (N)o : 
if %continue% == Y goto continue
if %continue% == y goto continue
cls
goto setfile
:continue
cls
:: create report export directory
md %filenum%-%exnum%
cd %filenum%-%exnum%
md Output
cd..
echo STARTING LIVE ACQUISITION...

:: Start Report File
echo TEA Version: %ver% > %filenum%-%exnum%/Evidence.txt
echo CASE FILE# %filenum%, Exhibit# %exnum%, Investigator: %invest% >> %filenum%-%exnum%/Evidence.txt
echo Application Description [Command run] ; Each output seperated by "_____" >> %filenum%-%exnum%/Evidence.txt
echo Commands run under user: %!!user% >> %filenum%-%exnum%/Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt


echo TIME STARTED [NOW] : >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/Time_S.txt
!!md5sum %filenum%-%exnum%/Output/Time_S.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\Time_S.txt >> %filenum%-%exnum%\Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt

::MEMORY DUMP
::FDPRO
if %mem%==m goto memd
if %mem%==M goto memd
if %mem%==p goto page
if %mem%==P goto page
::MOONSOLS
if %mem%==r goto rawmem
if %mem%==R goto rawmem
if %mem%==b goto biosmem
if %mem%==B goto biosmem
if %mem%==d goto devicsmem
if %mem%==D goto devicsmem
::##################WINENCAP?
if %mem%==c goto winen
if %mem%==C goto winen

goto osdetermine

::If memory dump selected
:memd
echo ==- MEMORY DUMP ONLY -==
echo ==- MEMORY DUMP ONLY -== >> %filenum%-%exnum%/Evidence.txt
echo TIME STARTED [NOW] : >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Evidence.txt
:: Set color (Red BG/Yellow Text)
color 4e
copy !!fdpro.exe %filenum%-%exnum%\Output
cd %filenum%-%exnum%/Output/
!!fdpro memdump.mem 
rename memdump.mem %filenum%-%exnum%.mem
del !!FDPro.exe
cd../..
echo TIME COMPLETED [NOW] : >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Evidence.txt
echo.
echo ^^^^PLEASE CHECK TO SEE IF FDPro WORKED. IF NOT, YOU DO NOT HAVE ADMIN RIGHTS.
echo Press control-c and run "admin-overide.cmd". If FDPro did work
pause
goto osdetermine

::If page file selected
echo ==- MEMORY DUMP WITH PAGE FILE -==
echo ==- MEMORY DUMP WITH PAGE FILE -== >> %filenum%-%exnum%/Evidence.txt
echo TIME STARTED [NOW] : >> %filenum%-%exnum%/Evidence.txt
:: Set color (Red BG/Yellow Text)
color 4e
:page
copy !!fdpro.exe %filenum%-%exnum%\Output
cd %filenum%-%exnum%/Output/
!!fdpro memdump.hpak
rename memdump.hpak %filenum%-%exnum%.hpak
del !!FDPro.exe
cd../..
echo TIME COMPLETED [NOW] : >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Evidence.txt
echo.
echo ^^^^PLEASE CHECK TO SEE IF FDPro WORKED. IF NOT, YOU DO NOT HAVE ADMIN RIGHTS.
echo Press control-c and run "admin-overide.cmd". If FDPro did work
pause
goto osdetermine


:rawmem
if %archi%==64 goto rawmem64

:rawmem32
echo ==- MOONSOLS RAW MEMORY DUMP -==
echo ==- MOONSOLS RAW MEMORY DUMP [win32dd.exe] -== >> %filenum%-%exnum%/Evidence.txt
echo TIME STARTED [NOW] : >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Evidence.txt
:: Set color (Red BG/Yellow Text)
color 4e
!!win32dd.exe /j %filenum%-%exnum%\MEMDUMP.log /f %filenum%-%exnum%\%filenum%-%exnum%.dmp
echo TIME COMPLETED [NOW] : >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Evidence.txt
echo.
goto osdetermine

:rawmem64
echo ==- MOONSOLS RAW MEMORY DUMP -==
echo ==- MOONSOLS RAW MEMORY DUMP [win64dd.exe] -== >> %filenum%-%exnum%/Evidence.txt
echo TIME STARTED [NOW] : >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Evidence.txt
:: Set color (Red BG/Yellow Text)
color 4e
!!win64dd.exe /j %filenum%-%exnum%\MEMDUMP.log /f %filenum%-%exnum%\%filenum%-%exnum%.dmp
echo TIME COMPLETED [NOW] : >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Evidence.txt
echo.
goto osdetermine

:biosmem
if %archi%==64 goto biosmem64
:biosmem32
echo ==- MOONSOLS RAW+BIOS MEMORY DUMP -==
echo ==- MOONSOLS RAW+BIOS MEMORY DUMP [win32dd.exe /r /c2] -== >> %filenum%-%exnum%/Evidence.txt
echo TIME STARTED [NOW] : >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Evidence.txt
:: Set color (Red BG/Yellow Text)
color 4e
!!win32dd.exe /r /c2 /j %filenum%-%exnum%\MEMDUMP.log /f %filenum%-%exnum%\%filenum%-%exnum%.dmp
echo TIME COMPLETED [NOW] : >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Evidence.txt
echo.
goto osdetermine

:biosmem64
echo ==- MOONSOLS RAW+BIOS MEMORY DUMP -==
echo ==- MOONSOLS RAW+BIOS MEMORY DUMP [win64dd.exe /r /c2] -== >> %filenum%-%exnum%/Evidence.txt
echo TIME STARTED [NOW] : >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Evidence.txt
:: Set color (Red BG/Yellow Text)
color 4e
!!win64dd.exe /r /c2 /j %filenum%-%exnum%\MEMDUMP.log /f %filenum%-%exnum%\%filenum%-%exnum%.dmp
echo TIME COMPLETED [NOW] : >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Evidence.txt
echo.
goto osdetermine


:devicsmem
if %archi%==64 goto devicsmem64
:devicsmem32
echo ==- MOONSOLS RAW+BIOS MEMORY DUMP -==
echo ==- MOONSOLS RAW+BIOS MEMORY DUMP [win32dd.exe /r /c0] -== >> %filenum%-%exnum%/Evidence.txt
echo TIME STARTED [NOW] : >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Evidence.txt
:: Set color (Red BG/Yellow Text)
color 4e
!!win32dd.exe /r /c0 /j %filenum%-%exnum%\MEMDUMP.log /f %filenum%-%exnum%\%filenum%-%exnum%.dmp
echo TIME COMPLETED [NOW] : >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Evidence.txt
echo.
goto osdetermine

:devicsmem64
echo ==- MOONSOLS RAW+BIOS MEMORY DUMP -==
echo ==- MOONSOLS RAW+BIOS MEMORY DUMP [win64dd.exe /r /c0] -== >> %filenum%-%exnum%/Evidence.txt
echo TIME STARTED [NOW] : >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Evidence.txt
:: Set color (Red BG/Yellow Text)
color 4e
!!win64dd.exe /r /c0 /j %filenum%-%exnum%\MEMDUMP.log /f %filenum%-%exnum%\%filenum%-%exnum%.dmp
echo TIME COMPLETED [NOW] : >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Evidence.txt
echo.
goto osdetermine

:winen
if %archi%==64 goto winen64
:winen32
echo Memory Capture [WinEn]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/WINEN.txt
!!winen -p "%filenum%-%exnum%/output/%filenum%-%exnum%" -m "%filenum%-%exnum%" -c "%filenum%" -e "%invest%" -r "%exnum%" -d 0 -n AcquiredwTEA
echo WINEN Started >> %filenum%-%exnum%/Output/WINEN.txt
!!md5sum %filenum%-%exnum%/Output/WINEN.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\WINEN.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto osdetermine

:winen64
echo Memory Capture [WinEn64]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/WINEN64.txt
!!winen64 -p "%filenum%-%exnum%/output/%filenum%-%exnum%" -m "%filenum%-%exnum%" -c "%filenum%" -e "%invest%" -r "%exnum%" -d 0 -n AcquiredwTEA
echo WINEN64 Started >> %filenum%-%exnum%/Output/WINEN64.txt
!!md5sum %filenum%-%exnum%/Output/WINEN64.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\WINEN64.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto osdetermine

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
::If not any of the above versions detected, hash the local tools against known Microsoft Hashes and run them if they match (windows 7 mode)
goto ver_Win7
::==========================================================================================

:blank
color e4
cls
echo FIELD CAN NOT BE LEFT BLANK!
pause
goto setfile

:ver_xp
:Run Windows XP specific commands here.
echo =========================================================================== >> %filenum%-%exnum%/Evidence.txt
echo ======================[Windows XP ARCHITECTURE DETECTED]======================= >> %filenum%-%exnum%/Evidence.txt
echo =============================[USB TOOLS ONLY]============================== >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
cls
echo ==- WINDOWS XP ARCHITECTURE DETECTED -==
:: Set color (Blue BG/White Text)
color 9f
echo In Progress...
goto run00

::run00
:xpip
echo NETWORK SETTINGS [IPCONFIG /ALL]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/IPCONFIG.txt
!!ipconfig /all >> %filenum%-%exnum%/Output/IPCONFIG.txt
!!md5sum %filenum%-%exnum%/Output/IPCONFIG.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\IPCONFIG.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run00a

::run00a
:xpdns
echo SHOW VISITED SITES [IPCONFIG /DISPLAYDNS]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/IPDNS.txt
!!ipconfig /displaydns >> %filenum%-%exnum%/Output/IPDNS.txt
!!md5sum %filenum%-%exnum%/Output/IPDNS.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\IPDNS.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run01

::run01
:xparp
echo MAC CONNECTIONS[ARP -A]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/ARP.txt
!!arp -a >> %filenum%-%exnum%/Output/ARP.txt
!!md5sum %filenum%-%exnum%/Output/ARP.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\ARP.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run02

::run02
:xproute
echo ROUTING TABLES [ROUTE PRINT]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/ROUTE.txt
echo. >> %filenum%-%exnum%/Output/ROUTE.txt
!!route print >> %filenum%-%exnum%/Output/ROUTE.txt
!!md5sum %filenum%-%exnum%/Output/ROUTE.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\ROUTE.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run03

::run03
:xpnetstat
echo PROTOCOL STATISTICS [NETSTAT -ABON]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/NETSTAT.txt
!!netstat -abon >> %filenum%-%exnum%/Output/NETSTAT.txt
!!md5sum %filenum%-%exnum%/Output/NETSTAT.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\NETSTAT.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run04

::run04
:xpnetses
echo OPEN SESSIONS [NET SESSIONS]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/NETSESN.txt
echo. >> %filenum%-%exnum%/Output/NETSESN.txt
!!net sessions >> %filenum%-%exnum%/Output/NETSESN.txt
!!md5sum %filenum%-%exnum%/Output/NETSESN.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\NETSESN.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run05

::run05
:xpnetview
echo DISCOVERABLE DEVICES ON NETWORK [NET VIEW]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/NETVIEW.txt
echo. >> %filenum%-%exnum%/Output/NETVIEW.txt
!!net view >> %filenum%-%exnum%/Output/NETVIEW.txt
!!md5sum %filenum%-%exnum%/Output/NETVIEW.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\NETVIEW.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run06

::run06
:xpnetshare
echo SHARES [NET SHARE]: >> %filenum%-%exnum%/Evidence.txt
echo THIS COULD TAKE A WHILE IF ON A NETWORK. IF IT HAS HUNG, PRESS CTRL-C AND THEN SELECT "N"
!!now >> %filenum%-%exnum%/Output/NETSHARE.txt
!!net share >> %filenum%-%exnum%/Output/NETSHARE.txt
!!md5sum %filenum%-%exnum%/Output/NETSHARE.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\NETSHARE.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run07

::run07
:xpat
echo DISPLAY SCHEDULED APPLICATIONS [AT]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/AT.txt
echo. >> %filenum%-%exnum%/Output/AT.txt
!!at >> %filenum%-%exnum%/Output/AT.txt
!!md5sum %filenum%-%exnum%/Output/AT.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\AT.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt

echo =========================================================================== >> %filenum%-%exnum%/Evidence.txt

goto both

::==========================================================================================
:ver_Win7
:Run Windows 7/Vista/Server 2008 specific commands here.
color 1e
cls
echo =========================================================================== >> %filenum%-%exnum%/Evidence.txt
echo =================[WINDOWS 7 / VISTA ARCHITECTURE DETECTED]================= >> %filenum%-%exnum%/Evidence.txt
echo [Tools (ipconfig, route, netstat, at, net, arp) run from %SystemRoot%/System32/] >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
echo ==- WINDOWS 7 / VISTA ARCHITECTURE DETECTED -==

goto cip 

:: HASH verification commented for ipconfig, other tools not commented.
::cip
:checkip
::Check IPCONFIG hash agaisnt Hash List
::Hash local version of ipconfig
!!md5sum %SystemRoot%/System32/ipconfig.exe > ipconfig.md5
::Set Varibles from md5 text file (this includes original file name and formating)
set /p ipconfig= < ipconfig.md5
:: delete temp file
del *.md5
:: Removes first character and reads 32 characters to obtain hash value.
set ipconfig=%ipconfig:~1,32%
::write hash to report
echo %SystemRoot%\System32\IPCONFIG.exe HASH: %ipconfig% >> %filenum%-%exnum%/Evidence.txt 
::display the hash
echo %SystemRoot%\System32\IPCONFIG.exe HASH: %ipconfig% 
::check hash against known list
find /c /i "%ipconfig%" !hashlist.tea > NUL
if %ERRORLEVEL% == 0 set ipc=1
if %ipc%==1 goto run10
echo IPCONFIG.exe HASH not found in known Database. Skipping... >> %filenum%-%exnum%/Evidence.txt
echo IPCONFIG.exe HASH not found in known Database. Skipping...
goto run11

::run10
:ipfound
if %ipc%==0 goto run11
echo IPCONFIG.exe Verified. Running application... >> %filenum%-%exnum%/Evidence.txt
echo IPCONFIG.exe verified.
echo Running application...
echo NETWORK SETTINGS [IPCONFIG /ALL]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/IPCONFIG.txt
ipconfig /all >> %filenum%-%exnum%/Output/IPCONFIG.txt
!!md5sum %filenum%-%exnum%/Output/IPCONFIG.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\IPCONFIG.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
echo.
goto run10a

::run10a
:checkdns
::NEED TO CHECK IP
if %ipc%==0 goto run11
echo IPCONFIG.exe Verified. Running application... >> %filenum%-%exnum%/Evidence.txt
echo IPCONFIG.exe [/displaydns] verified.
echo Running application...
echo SHOW SITES VISITED [IPCONFIG /DISPLAYDNS]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/IPDNS.txt
ipconfig /displaydns >> %filenum%-%exnum%/Output/IPDNS.txt
!!md5sum %filenum%-%exnum%/Output/IPDNS.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\IPDNS.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
echo.
goto run11

::run11
:checkat
!!md5sum %SystemRoot%/System32/at.exe > at.md5
set /p at= < at.md5
del *.md5
echo %SystemRoot%\System32\AT.exe HASH: %at% >> %filenum%-%exnum%/Evidence.txt 
echo %SystemRoot%\System32\AT.exe HASH: %at% 
find /c /i "%at%" !hashlist.tea > NUL
if %ERRORLEVEL% == 0 goto atfound
echo AT.exe HASH not found in known Database. Skipping... >> %filenum%-%exnum%/Evidence.txt
echo AT.exe HASH not found in known Database. Skipping...
goto run12

:atfound
echo AT.exe Verified. Running application... >> %filenum%-%exnum%/Evidence.txt
echo AT.exe Verified.
echo Running application...
echo DISPLAY SCHEDULED APPLICATIONS [AT]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/AT.txt
echo. >> %filenum%-%exnum%/Output/AT.txt
at >> %filenum%-%exnum%/Output/AT.txt
!!md5sum %filenum%-%exnum%/Output/AT.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\AT.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
echo.
goto run12

::run12
:checkroute
!!md5sum %SystemRoot%/System32/route.exe > route.md5
set /p route= < route.md5
del *.md5
set route=%route:~1,32%
echo %SystemRoot%\System32\ROUTE.exe HASH: %route% >> %filenum%-%exnum%/Evidence.txt 
echo %SystemRoot%\System32\ROUTE.exe HASH: %route% 
find /c /i "%route%" !hashlist.tea > NUL
if %ERRORLEVEL% == 0 goto routefound
echo ROUTE.exe HASH not found in known Database. Skipping... >> %filenum%-%exnum%/Evidence.txt
echo ROUTE.exe HASH not found in known Database. Skipping... 
goto run13
:routefound
echo route.exe Verified. Running application... >> %filenum%-%exnum%/Evidence.txt
echo route.exe Verified.
echo Running application... 
echo ROUTING TABLES [ROUTE PRINT]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/ROUTE.txt
echo. >> %filenum%-%exnum%/Output/ROUTE.txt
route print >> %filenum%-%exnum%/Output/ROUTE.txt
!!md5sum %filenum%-%exnum%/Output/ROUTE.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\ROUTE.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
echo.
goto run13

::run13
:checknetstat
!!md5sum %SystemRoot%/System32/netstat.exe > netstat.md5
set /p netstat= < netstat.md5
del *.md5
set netstat=%netstat:~1,32%
echo %SystemRoot%\System32\NETSTAT.exe HASH: %netstat% >> %filenum%-%exnum%/Evidence.txt 
echo %SystemRoot%\System32\NETSTAT.exe HASH: %netstat% 
find /c /i "%netstat%" !hashlist.tea > NUL
if %ERRORLEVEL% == 0 goto netstatfound
echo NETSTAT.exe HASH not found in known Database. Skipping... >> %filenum%-%exnum%/Evidence.txt
echo NETSTAT.exe HASH not found in known Database. Skipping... 
goto run14
:netstatfound
echo netstat.exe Verified. Running application... >> %filenum%-%exnum%/Evidence.txt
echo netstat.exe Verified.
echo Running application... 
echo PROTOCOL STATISTICS [NETSTAT -ABON]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/NETSTAT.txt
netstat -abon >> %filenum%-%exnum%/Output/NETSTAT.txt
!!md5sum %filenum%-%exnum%/Output/NETSTAT.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\NETSTAT.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
echo.
goto cnet

::cnet
:checknet
!!md5sum %SystemRoot%/System32/net.exe > net.md5
set /p net= < net.md5
del *.md5
set net=%net:~1,32%
echo %SystemRoot%\System32\NET.exe HASH: %net%  >> %filenum%-%exnum%/Evidence.txt
echo %SystemRoot%\System32\NET.exe HASH: %net%
find /c /i "%net%" !hashlist.tea > NUL
if %ERRORLEVEL% == 0 set cnet=1
if %cnet%==1 goto run14
echo net.exe HASH not found in known Database. Skipping... >> %filenum%-%exnum%/Evidence.txt
echo net.exe HASH not found in known Database. Skipping... 
goto run17
::run14
:netfound
if %cnet%==0 goto run17
echo net.exe Verified. Running application... >> %filenum%-%exnum%/Evidence.txt
echo net.exe Verified.
echo Running application... 
echo OPEN SESSIONS [NET SESSIONS]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/NETSESN.txt
echo. >> %filenum%-%exnum%/Output/NETSESN.txt
net sessions >> %filenum%-%exnum%/Output/NETSESN.txt
!!md5sum %filenum%-%exnum%/Output/NETSESN.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\NETSESN.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run15

::run15
:7netview
if %cnet%==0 goto run17
echo DISCOVERABLE DEVICES ON NETWORK [NET VIEW]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/NETVIEW.txt
echo. >> %filenum%-%exnum%/Output/NETVIEW.txt
net view >> %filenum%-%exnum%/Output/NETVIEW.txt
!!md5sum %filenum%-%exnum%/Output/NETVIEW.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\NETVIEW.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run16

::run16
:7netshare
if %cnet%==0 goto run17
echo SHARES [NET SHARE]: >> %filenum%-%exnum%/Evidence.txt
echo THIS COULD TAKE A WHILE IF ON A NETWORK. IF IT HAS HUNG, TYPE "N" - ENTER". IF NO SUCCESS, PRESS CTRL-C AND THEN SELECT "N".
!!now >> %filenum%-%exnum%/Output/NETSHARE.txt
net share >> %filenum%-%exnum%/Output/NETSHARE.txt
!!md5sum %filenum%-%exnum%/Output/NETSHARE.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\NETSHARE.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
echo.
goto run17

::run17
:checkarp
!!md5sum %SystemRoot%/System32/arp.exe > arp.md5
set /p arp= < arp.md5
del *.md5
set arp=%arp:~1,32%
echo %SystemRoot%\System32\ARP.exe HASH:%arp%  >> %filenum%-%exnum%/Evidence.txt
echo %SystemRoot%\System32\ARP.exe HASH:%arp%
::Check arp hash agaisnt Hash List
find /c /i "%arp%" !hashlist.tea > NUL
if %ERRORLEVEL% == 0 goto arpfound
echo arp.exe HASH not found in known Database. Skipping... >> %filenum%-%exnum%/Evidence.txt
echo arp.exe HASH not found in known Database. Skipping...
goto both
:arpfound
echo arp.exe Verified. Running application... >> %filenum%-%exnum%/Evidence.txt
echo arp.exe Verified.
echo Running application... 
echo MAC CONNECTIONS[ARP -A]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/ARP.txt
arp -a >> %filenum%-%exnum%/Output/ARP.txt
!!md5sum %filenum%-%exnum%/Output/ARP.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\ARP.txt >> %filenum%-%exnum%\Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
echo.

goto both

:both
:cls
echo =========================================================================== >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
echo ====================[USB TOOLS FOR ALL WINDOWS VERSIONS]=================== >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
echo ==- USB TOOLS FOR ALL WINDOWS VERSIONS -==
echo In Progress...
:: Set color (Black BG/Green Text)
color 0a
goto run20
::run20
:rnnbtstat
echo NETBIOS NAMES [NBTSTAT -N]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/NBTSTAT.txt
!!nbtstat -n >> %filenum%-%exnum%/Output/NBTSTAT.txt
!!md5sum %filenum%-%exnum%/Output/NBTSTAT.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\NBTSTAT.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run21

::run21
:rnopenports
echo OPEN PORTS [OPENPORTS -LINES -PATH] >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/OPENPORT.txt
echo. >> %filenum%-%exnum%/Output/OPENPORT.txt
!!openports -lines -path >> %filenum%-%exnum%/Output/OPENPORT.txt
!!md5sum %filenum%-%exnum%/Output/OPENPORT.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\OPENPORT.txt >> %filenum%-%exnum%\Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run22

::run22
:rnpsloggedon
echo USERS LOGGED ON [PSLOGGEDON /ACCEPTEULA]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/PSLOGGED.txt
echo. >> %filenum%-%exnum%/Output/PSLOGGED.txt
!!psloggedon /accepteula >> %filenum%-%exnum%/Output/PSLOGGED.txt
!!md5sum %filenum%-%exnum%/Output/PSLOGGED.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\PSLOGGED.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run23

::run23
:rnpslist
echo PROCESSES RUNNING [PSLIST -T]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/PSLIST.txt
echo. >> %filenum%-%exnum%/Output/PSLIST.txt
echo Abbreviation key: >> %filenum%-%exnum%/Output/PSLIST.txt
echo    Pri         Priority >> %filenum%-%exnum%/Output/PSLIST.txt
echo    Thd         Number of Threads >> %filenum%-%exnum%/Output/PSLIST.txt
echo    Hnd         Number of Handles >> %filenum%-%exnum%/Output/PSLIST.txt
echo    VM          Virtual Memory >> %filenum%-%exnum%/Output/PSLIST.txt
echo    WS          Working Set >> %filenum%-%exnum%/Output/PSLIST.txt
echo    Priv        Private Virtual Mem >> %filenum%-%exnum%/Output/PSLIST.txt
echo    Priv Pk     Private Virtual Mem >> %filenum%-%exnum%/Output/PSLIST.txt
echo    Faults      Page Faults >> %filenum%-%exnum%/Output/PSLIST.txt
echo    NonP        Non-Paged Pool >> %filenum%-%exnum%/Output/PSLIST.txt
echo    Page        Paged Pool >> %filenum%-%exnum%/Output/PSLIST.txt
echo    Cswtch      Context Switches >> %filenum%-%exnum%/Output/PSLIST.txt
echo  --------------------------------  >> %filenum%-%exnum%/Output/PSLIST.txt
echo.  >> %filenum%-%exnum%/Output/PSLIST.txt
!!pslist -t /accepteula >> %filenum%-%exnum%/Output/PSLIST.txt
!!md5sum %filenum%-%exnum%/Output/PSLIST.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\PSLIST.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run24

::run24
:rnpsfile
echo OPEN FILES [PSFILE /ACCEPTEULA]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/PSFILE.txt
echo. >> %filenum%-%exnum%/Output/PSFILE.txt
!!psfile /accepteula >> %filenum%-%exnum%/Output/PSFILE.txt
!!md5sum %filenum%-%exnum%/Output/PSFILE.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\PSFILE.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run25

::run25
:rnsysinfo
echo SYSTEM INFO [SYSTEMINFO]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/SYSINFO.txt
!!Systeminfo >> %filenum%-%exnum%/Output/SYSINFO.txt
!!md5sum %filenum%-%exnum%/Output/SYSINFO.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\SYSINFO.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run26

::run26
:rnpsinfo
echo SYSTEM PATCHES [PSINFO -H /ACCEPTEULA]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/PSinfo.txt
echo. >> %filenum%-%exnum%/Output/PSinfo.txt
!!psinfo -h /accepteula >> %filenum%-%exnum%/Output/PSinfo.txt
!!md5sum %filenum%-%exnum%/Output/PSinfo.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\PSinfo.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
goto run27

::run27
:rnwhoami
echo LOGGED IN USER AND SID [WHOAMI /USER /SID]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/WhoamI.txt
echo. >> %filenum%-%exnum%/Output/WhoamI.txt
!!whoami /user /sid >> %filenum%-%exnum%/Output/WhoamI.txt
!!md5sum %filenum%-%exnum%/Output/WhoamI.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\WhoamI.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run28

::run28
:rngetmac
echo DISPLAY PHYSICAL / MAC ADDRESS [GETMAC /V]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/MAC.txt
!!getmac /v >> %filenum%-%exnum%/Output/MAC.txt
!!md5sum %filenum%-%exnum%/Output/MAC.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\MAC.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run29

::run29
:rncurl
echo IP ADDRESS SEEN EXTERNALY [CURL]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/CURL.txt
echo. >> %filenum%-%exnum%/Output/CURL.txt
!!curl http://www.showmyip.com/simple/ >> %filenum%-%exnum%/Output/CURL.txt
!!md5sum %filenum%-%exnum%/Output/CURL.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\CURL.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
goto run29B


::CHECK FOR ENCRYPTED VOLUMES
::run29B
:rndriveid
::CHECK FOR DRIVEID
IF EXIST !!driveid.exe goto driveid
echo DRIVE ID Application not detected, running EDD...
echo Check for Encryption [EDD /Batch]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/EDD.txt
echo. >> %filenum%-%exnum%/Output/EDD.txt
!!edd /batch >> %filenum%-%exnum%/Output/EDD.txt
!!md5sum %filenum%-%exnum%/Output/EDD.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\EDD.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
cd %filenum%-%exnum%/Output
cls
type EDD.txt
cd ../..
echo PLEASE REVIEW ENCRYPTION STATUS. IF ENCRYPTION MOUNTED, CONSIDER LOGICAL ACQUISITION.
goto run30

:driveid
echo DRIVE ID [ENCRYPTION DETECTION]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/DRIVEID.txt
echo. >> %filenum%-%exnum%/Output/DRIVEID.txt
!!driveid >> %filenum%-%exnum%/Output/DRIVEID.txt
!!md5sum %filenum%-%exnum%/Output/DRIVEID.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\DRIVEID.txt >> %filenum%-%exnum%\Evidence.txt
echo ___________________________________________________________________________ >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
cd %filenum%-%exnum%/Output
type DRIVEID.txt
cd ../..
echo.
echo.
echo PLEASE REVIEW ENCRYPTION STATUS!! IF ENCRYPTION MOUNTED, CONSIDER LOGICAL ACQUISITION.
goto run30

:endfile
echo =========================================================================== >> %filenum%-%exnum%/Evidence.txt
echo TIME COMPLETED [NOW]: >> %filenum%-%exnum%/Evidence.txt
!!now >> %filenum%-%exnum%/Output/Time_E.txt
!!md5sum %filenum%-%exnum%/Output/Time_E.txt >> %filenum%-%exnum%/Output/HASH_VALUES.txt
type %filenum%-%exnum%\Output\Time_E.txt >> %filenum%-%exnum%\Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
echo =========================================================================== >> %filenum%-%exnum%/Evidence.txt

echo HASH_VALUES.txt MD5 HASH: >> %filenum%-%exnum%/Evidence.txt
echo. >> %filenum%-%exnum%/Evidence.txt
!!md5sum %filenum%-%exnum%/Output/HASH_VALUES.txt >> %filenum%-%exnum%/Evidence.txt
goto exit

:admininstr
cls
color 0c
echo.
echo.
echo To run command prompt as an Administrator
echo in Windows 7,Press the windows key, type
echo "cmd" and then press control-shift-enter.
echo.
echo.
echo.
echo ==================================================
echo ==== You can also run as SYSTEM by            ====
echo ==== running SYSTEM.CMD from the new command  ====
echo ==== Prompt which should bypass any other     ====
echo ==== Privledge limitations.                   ==== 
echo ==================================================
echo.
echo.
echo.
echo Recommend press control-c to cancel or to continue 
pause
goto start
:brew
echo.
echo.
echo NO CONFIG FILE FOUND, please run BREW 
echo (Recommended running on analysis machine not target!)
:: NOTE: running brew on target box will leave artifacts in memory.
echo.
echo.
echo.
goto end
:exit
:: reset color
echo.
echo.
echo.
echo Live Acquisition complete.
echo Wait 15 seconds then unplug power or remove battery.
echo In cases where Memory Capture failed type: HIBERNATE
echo WARNING - HIBERNATE WILL MAKE CHANGES TO THE HARD DRIVE 
echo **Windows will dump memory to the HDD and overwrite unallocated space**
echo **This 'could' result in the destruction of evidence and must be **
echo **assessed on a case by case basis.**
COLOR 07
echo.
echo.
echo.
echo.
pause
:end