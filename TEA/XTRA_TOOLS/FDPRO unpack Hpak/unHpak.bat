@echo off
cls
:: UnHPAK Batch
:: ######################################################
:: VERSION 0.1.110321
:: SCRIPT: UnHPAK Batch
:: CREATION DATE: 2011-03-21
:: LAST MODIFIED: 2011-03-23
:: AUTHOR: Cst. Mark SOUTHBY, #53384
:: EMAIL: mark.southby@rcmp-grc.gc.ca
:: ######################################################
:: DESCRIPTION: Simple batch to extract Memory dump and page file from FDPro memory dumps.
:: ######################################################

::Check for FDPRo.exe
:FDProchk
IF EXIST !!FDPRo.exe goto begin
color 0c
echo !!FDPro.exe NOT FOUND. PLEASE PLACE THIS BATCH IN YOUR TEA FOLDER or rename FDPro.exe to !!FDPro.exe 
pause
goto exit

:begin
::grab path1 and file name. Click
set /p path1=Type full path to memory file [Click and drag file here from explorer and press enter]:
set /p exnum=File name for output files [file-exhibit]:

echo Extracting data from memory dump... 
echo To skip a command, press CTRL-C and select (N) to terminate batch.

echo Extracting Memory to %exnum%memdump.bin
!!fdpro %path1% -hpak extract 0
rename memdump.bin %exnum%memdump.bin

echo Extracting Pagefile to %exnum%dumpfile.sys
!!fdpro %path1% -hpak extract 1
rename dumpfile.sys %exnum%dumpfile.sys

color 0c
echo PLEASE NOTE: The extracted files, %exnum%memdump.bin and %exnum%dumpfile.sys are in the root folder.
echo Please move them to the same folder as %path1%.
pause

:exit
color 0f
::NFAR CH
::ms-53384
