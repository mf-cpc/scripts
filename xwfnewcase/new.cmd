@echo off
setlocal
set CURDIR=%~dp0
set SRCPATH=%~dp0
set CASEFOLDERSTRUCTURE="C:\CASE_FOLDER_STRUCTURE"
set DRIVE=C
if exist %CASEFOLDERSTRUCTURE% goto continue
echo %CASEFOLDERSTRUCTURE% not found. Can't continue.
exit /b

:continue
echo %CASEFOLDER% exist
echo. 
set /p "DRIVE=Enter drive letter where to create new case or just ENTER for default [%DRIVE%] : "
set /p "CASENUMBER=Enter Case Number/Name: " 
::echo Creating %DRIVE%:\%CASENUMBER%
::%DRIVE%:
::cd\
echo Copy in progress..... 
robocopy /MIR /NFL /NDL /NJH /NJS /ETA %CASEFOLDERSTRUCTURE% %DRIVE%:\%CASENUMBER%
echo Copy completed
echo Creating desktop shortcut

: Set the target file or command for the shortcut
set "TARGETFILE=%DRIVE%:\%CASENUMBER%\xwf\xwforensics64.exe"
: Set the path for the desktop
set "DESKTOPPATH=%userprofile%\Desktop"
: Set the name for the shortcut
set "SHORTCUTNAME=%CASENUMBER%.lnk"

echo %TARGETFILE%
echo %DESKTOPPATH%
echo %SHORTCUTNAME%
break
:Create the shortcut

echo set WSHSHELL = wscript.CreateObject("WScript.Shell") > "%desktopPath%\%SHORTCUTNAME%"
echo set SHORTCUT = WSHSHELL.CreateShortcut("%DESKTOPPATH%\%SHORTCUTNAME%") >> "%DESKTOPPATH%\%SHORTCUTNAME%"
echo SHORTCUT.TargetPath = "%TARGETFILE%" >> "%DESKTOPPATH%\%SHORTCUTNAME%"
echo SHORTCUT.Save >> "%DESKTOPPATH%\%SHORTCUTNAME%"
echo Shortcut created: %DESKTOPPATH%\%SHORTCUTNAME%
endlocal