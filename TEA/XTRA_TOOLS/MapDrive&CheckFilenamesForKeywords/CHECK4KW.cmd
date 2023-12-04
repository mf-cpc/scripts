@echo off
:: ######################################################
:: VERSION 0.9.121221a - December 21 2012 Apocolypse Edition
:: SCRIPT: Check 4 Key Words
:: CREATION DATE: 2012-12-21
:: LAST MODIFIED: 2012-12-21
:: Hash list current as of: 
:: AUTHOR: Cst. Mark SOUTHBY, #53384
:: EMAIL: mark.southby@rcmp-grc.gc.ca
:: ######################################################
:: DESCRIPTION: A batch file created to run run keyword searches on live machines. Edit key_words.tea to change keywords.
:: ######################################################
cls
echo listing drives ....
wmic logicaldisk list brief
echo. 
echo.
:map
set drive=0
echo Leave drive letter blank (hit enter) once all desired drives mapped.
echo.
echo.
echo.
echo ***Only put letter in lower case not semicolon (eg: c for C:)!!! ***
echo.
set /p drive=Drive Letter to map:
if %drive% == 0 goto mapped
echo Mapping Drive %drive%:, this may take a while please wait...
dir %drive%:\ /b /s > %drive%_dir.txt
echo Map for drive %drive%: Generated.
goto map
:mapped
echo Searching drive maps for filenames containing related keywords...
findstr /G:key_words.tea *_dir.txt >KWHits.txt
:: todo: remove the first 10 characters of KWHits.txt
echo Total number of hits:
find /c /v "" KWHits.txt
echo.
Echo Would you like to view the locations of the files?
set /p yno=(Y)es or (N)o:
if %yno% == y type KWHits.txt|more
echo.
echo Please note the following files were generated:
dir *.txt /b
echo.
echo. As these contain the entire file listings for that drive, you must delete these file manually.
echo.
echo.
echo.
echo Completed.
echo To end batch
pause
:exit