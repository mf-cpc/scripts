@echo off
cls
:: ######################################################
:: VERSION 1.1.211214C - CMPFOR Edition
:: SCRIPT: TEA TOOL VERIFICATION
:: CREATION DATE: 2011-02-01
:: LAST MODIFIED: 2021-12-14
:: AUTHOR: Mark SOUTHBY, #53384
:: EMAIL: mark.southby@rcmp-grc.gc.ca
:: ######################################################
:: #### For Law enforcement use only. Can be used and modified without persmission for Law Enforcement purposes.
:: 
:: HASH AND VERIFY REPORT
::#########################
IF EXIST *.td5 del *.td5
set err=0
::get folder
IF [%1] == [] GOTO set1
set path1=%1
goto TER
:set1
set /p path1=Click and drag folder from explorer to here and press enter: 
cls
echo Validating Report integrity...
echo.
:: ==============================================
:TER
echo Checking TEA Evidence Report integrity
dir /b %path1%\evidence*txt > $TER.td5
set /p TER= < $TER.td5
set TER=%TER:~-12,-4%
!!md5sum %path1%\Evidence*txt >$TERm.td5
set /p TERm= < $TERm.td5
set TERm=%TERm:~8,8%
if %TER%==%TERM% goto TERf
color 4E
set err=1
echo Evidence.%TER%.txt Verification failed!!
echo.
pause
goto step3
:TERf
echo TEA Evidence Report Verified!
echo.
goto step1
:: ==============================================
:step1
::find hash of hash_values.txt
type %path1%\Evidence.%TER%.txt |findstr Output > $ev_hash.td5
set /p ev_hash= < $ev_hash.td5
set ev_hash=%ev_hash:~1,32%
echo Report: %ev_hash%
::Get Actual hash of HASH_VALUES.txt
!!md5sum %path1%\Output\HASH_VALUES.txt >$hash_values.td5
set /p hash_values= < $hash_values.td5
set hash_values=%hash_values:~1,32%
echo Actual: %hash_values%
::Compare the two
if %ev_hash%==%hash_values% goto step2
:mismatch
color 0c
set err=1
echo The current HASH_VALUES.txt does not match the Hash reported!!
echo. 
echo Current HASH_VALUES.txt hash: %hash_values%
echo Hash Reported in report: %ev_hash%
echo.
pause
goto step3

:step2
color 0a
echo HASH_VALUES.txt not altered.
echo.
::pause
goto step3

:step3
:: ================ Check individual TXT files ================
echo ============================================
echo Checking integrity of original Output files:
echo.
:: Time Start
:step4
:ts
IF EXIST %path1%\OUTPUT\Time_S.txt goto time
:we
IF EXIST %path1%\OUTPUT\WINEN.txt goto winen
:we64
IF EXIST %path1%\OUTPUT\WINEN64.txt goto winen64
:ip
IF EXIST %path1%\OUTPUT\IPCONFIG.txt goto ipconfig
:ipd
IF EXIST %path1%\OUTPUT\ipdns.txt goto ipdns
:ar
IF EXIST %path1%\OUTPUT\arp.txt goto arp
:rt
IF EXIST %path1%\OUTPUT\route.txt goto route
:ns
IF EXIST %path1%\OUTPUT\netstat.txt goto netstat
:nss
IF EXIST %path1%\OUTPUT\netsesn.txt goto netsesn
:nv
IF EXIST %path1%\OUTPUT\netview.txt goto netview
:nsh
IF EXIST %path1%\OUTPUT\netshare.txt goto netshare
:a
IF EXIST %path1%\OUTPUT\at.txt goto at
:nbt
IF EXIST %path1%\OUTPUT\nbtstat.txt goto nbtstat
:op
IF EXIST %path1%\OUTPUT\openport.txt goto openport
:plog
IF EXIST %path1%\OUTPUT\pslogged.txt goto pslogged
:plst
IF EXIST %path1%\OUTPUT\pslist.txt goto pslist
:pf
IF EXIST %path1%\OUTPUT\psfile.txt goto psfile
:si
IF EXIST %path1%\OUTPUT\sysinfo.txt goto sysinfo
:pi
IF EXIST %path1%\OUTPUT\psinfo.txt goto psinfo
:wi
IF EXIST %path1%\OUTPUT\whoami.txt goto whoami
:mc
IF EXIST %path1%\OUTPUT\mac.txt goto mac
:cl
IF EXIST %path1%\OUTPUT\curl.txt goto curl
:ed
IF EXIST %path1%\OUTPUT\edd.txt goto edd
:di
IF EXIST %path1%\OUTPUT\driveid.txt goto driveid
:te
IF EXIST %path1%\OUTPUT\time_e.txt goto time_e
goto exit
:: If file exists, hash it to file $__.md5

:time
:: ==============================================
:: hash the report file
echo Time_S.txt Found, checking...
!!md5sum %path1%\OUTPUT\Time_S.txt > $time_s.td5
:: set the hash as a variable  
set /p time_s= < $time_s.td5
set time_s=%time_s:~1,32%
echo Report: %time_s%
:: Grab the hash from hashlist.txt
type %path1%\OUTPUT\HASH_VALUES.txt |findstr Time_S.txt > $ts.td5
set /p ts= < $ts.td5
set ts=%ts:~1,32%
echo Actual: %ts%
:: compare the two
if %ts%==%time_s% goto tf
color 0c
set err=1
echo Time_S.txt Hash Mismatch!!
echo.
goto we
:tf
echo Verified!
echo.
goto we
:: ==============================================


:winen
echo WINEN.txt Found, checking...
!!md5sum %path1%\OUTPUT\WINEN.txt > $WINEN.td5
set /p WINEN= < $WINEN.td5
set WINEN=%WINEN:~1,32%
echo Report: %WINEN%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr WINEN.txt > $we.td5
set /p we= < $we.td5
set we=%we:~1,32%
echo Actual: %we%
if %we%==%WINEN% goto wf
color 0c
set err=1
echo WINEN.txt Hash Mismatch!!
echo.
goto we64
:wf
echo Verified!
echo.
goto we64
:: ==============================================
:winen64
echo WINEN64.txt Found, checking...
!!md5sum %path1%\OUTPUT\WINEN64.txt > $WINEN64.td5
set /p WINEN64= < $WINEN64.td5
set WINEN64=%WINEN64:~1,32%
echo Report: %WINEN64%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr WINEN64.txt > $we.td5
set /p we64= < $we64.td5
set we64=%we64:~1,32%
echo Actual: %we64%
if %we64%==%WINEN64% goto w64f
color 0c
echo WINEN64.txt Hash Mismatch!!
echo.
goto ip
:w64f
echo Verified!
echo.
goto ip
:: ==============================================
:ipconfig
echo IPCONFIG.txt Found, checking...
!!md5sum %path1%\OUTPUT\IPCONFIG.txt > $IPCONFIG.td5
set /p IPCONFIG= < $IPCONFIG.td5
set IPCONFIG=%IPCONFIG:~1,32%
echo Report: %IPCONFIG%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr IPCONFIG.txt > $ip.td5
set /p ip= < $ip.td5
set ip=%ip:~1,32%
echo Actual: %ip%
if %ip%==%IPCONFIG% goto ipf
color 0c
set err=1
echo IPCONFIG.txt Hash Mismatch!!
echo.
pause
goto ipd
:ipf
echo Verified!
echo.
goto ipd
:: ==============================================
:ipdns
echo IPDNS.txt Found, checking...
!!md5sum %path1%\OUTPUT\IPDNS.txt > $IPDNS.td5
set /p IPDNS= < $IPDNS.td5
set IPDNS=%IPDNS:~1,32%
echo Report: %IPDNS%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr IPDNS.txt > $ipd.td5
set /p ipd= < $ipd.td5
set ipd=%ipd:~1,32%
echo Actual: %ipd%
if %ipd%==%IPDNS% goto ipdf
color 0c
set err=1
echo IPDNS.txt Hash Mismatch!!
echo.
pause
goto ar
:ipdf
echo Verified!
echo.
goto ar
:: ==============================================
:arp
echo ARP.txt Found, checking...
!!md5sum %path1%\OUTPUT\ARP.txt > $ARP.td5
set /p ARP= < $ARP.td5
set ARP=%ARP:~1,32%
echo Report: %ARP%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr ARP.txt > $ar.td5
set /p ar= < $ar.td5
set ar=%ar:~1,32%
echo Actual: %ar%
if %ar%==%ARP% goto arf
color 0c
set err=1
echo ARP.txt Hash Mismatch!!
echo.
pause
goto rt
:arf
echo Verified!
echo.
goto rt
:: ==============================================
:route
echo ROUTE.txt Found, checking...
!!md5sum %path1%\OUTPUT\ROUTE.txt > $ROUTE.td5
set /p ROUTE= < $ROUTE.td5
set ROUTE=%ROUTE:~1,32%
echo Report: %ROUTE%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr ROUTE.txt > $rt.td5
set /p rt= < $rt.td5
set rt=%rt:~1,32%
echo Actual: %rt%
if %rt%==%ROUTE% goto rf
color 0c
set err=1
echo ROUTE.txt Hash Mismatch!!
echo.
pause
goto ns
:rf
echo Verified!
echo.
goto ns
:: ==============================================
:netstat
echo NETSTAT.txt Found, checking...
!!md5sum %path1%\OUTPUT\NETSTAT.txt > $NETSTAT.td5
set /p NETSTAT= < $NETSTAT.td5
set NETSTAT=%NETSTAT:~1,32%
echo Report: %NETSTAT%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr NETSTAT.txt > $ns.td5
set /p ns= < $ns.td5
set ns=%ns:~1,32%
echo Actual: %ns%
if %ns%==%NETSTAT% goto nsf
color 0c
set err=1
echo NETSTAT.txt Hash Mismatch!!
echo.
pause
goto nss
:nsf
echo Verified!
echo.
goto nss
:: ==============================================
:netsesn
echo NETSESN.txt Found, checking...
!!md5sum %path1%\OUTPUT\NETSESN.txt > $NETSESN.td5
set /p NETSESN= < $NETSESN.td5
set NETSESN=%NETSESN:~1,32%
echo Report: %NETSESN%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr NETSESN.txt > $nss.td5
set /p nss= < $nss.td5
set nss=%nss:~1,32%
echo Actual: %nss%
if %nss%==%NETSESN% goto nssf
color 0c
set err=1
echo NETSESN.txt Hash Mismatch!!
echo.
pause
goto nv
:nssf
echo Verified!
echo.
goto nv
:: ==============================================
:netview
echo NETVIEW.txt Found, checking...
!!md5sum %path1%\OUTPUT\NETVIEW.txt > $NETVIEW.td5
set /p NETVIEW= < $NETVIEW.td5
set NETVIEW=%NETVIEW:~1,32%
echo Report: %NETVIEW%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr NETVIEW.txt > $nv.td5
set /p nv= < $nv.td5
set nv=%nv:~1,32%
echo Actual: %nv%
if %nv%==%NETVIEW% goto nvf
color 0c
set err=1
echo NETVIEW.txt Hash Mismatch!!
echo.
pause
goto nsh
:nvf
echo Verified!
echo.
goto nsh
:: ==============================================
:netshare
echo NETSHARE.txt Found, checking...
!!md5sum %path1%\OUTPUT\NETSHARE.txt > $NETSHARE.td5
set /p NETSHARE= < $NETSHARE.td5
set NETSHARE=%NETSHARE:~1,32%
echo Report: %NETSHARE%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr NETSHARE.txt > $nsh.td5
set /p nsh= < $nsh.td5
set nsh=%nsh:~1,32%
echo Actual: %nsh%
if %nsh%==%NETSHARE% goto nshf
color 0c
set err=1
echo NETSHARE.txt Hash Mismatch!!
echo.
pause
goto nsh
:nshf
echo Verified!
echo.
goto a
:: ==============================================
:at
echo AT.txt Found, checking...
!!md5sum %path1%\OUTPUT\AT.txt > $AT.td5
set /p AT= < $AT.td5
set AT=%AT:~1,32%
echo Report: %AT%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr AT.txt > $a.td5
set /p a= < $a.td5
set a=%a:~1,32%
echo Actual: %a%
if %a%==%AT% goto atf
color 0c
set err=1
echo AT.txt Hash Mismatch!!
echo.
pause
goto nbt
:atf
echo Verified!
echo.
goto nbt
:: ==============================================
:nbtstat
echo NBTSTAT.txt Found, checking...
!!md5sum %path1%\OUTPUT\NBTSTAT.txt > $NBTSTAT.td5
set /p NBTSTAT= < $NBTSTAT.td5
set NBTSTAT=%NBTSTAT:~1,32%
echo Report: %NBTSTAT%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr NBTSTAT.txt > $nbt.td5
set /p nbt= < $nbt.td5
set nbt=%nbt:~1,32%
echo Actual: %nbt%
if %nbt%==%NBTSTAT% goto nbtf
color 0c
set err=1
echo NBTSTAT.txt Hash Mismatch!!
echo.
pause
goto op
:nbtf
echo Verified!
echo.
goto op
:: ==============================================
:openport
echo OPENPORT.txt Found, checking...
!!md5sum %path1%\OUTPUT\OPENPORT.txt > $OPENPORT.td5
set /p OPENPORT= < $OPENPORT.td5
set OPENPORT=%OPENPORT:~1,32%
echo Report: %OPENPORT%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr OPENPORT.txt > $opp.td5
set /p opp= < $opp.td5
set opp=%opp:~1,32%
echo Actual: %opp%
if %opp%==%OPENPORT% goto oppf
color 0c
set err=1
echo OPENPORT.txt Hash Mismatch!!
echo.
pause
goto pslogged
:oppf
echo Verified!
echo.
goto pslogged
:: ==============================================
:pslogged
echo PSLOGGED.txt Found, checking...
!!md5sum %path1%\OUTPUT\PSLOGGED.txt > $PSLOGGED.td5
set /p PSLOGGED= < $PSLOGGED.td5
set PSLOGGED=%PSLOGGED:~1,32%
echo Report: %PSLOGGED%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr PSLOGGED.txt > $psl.td5
set /p psl= < $psl.td5
set psl=%psl:~1,32%
echo Actual: %psl%
if %psl%==%PSLOGGED% goto pslf
color 0c
set err=1
echo PSLOGGED.txt Hash Mismatch!!
echo.
pause
goto plst
:pslf
echo Verified!
echo.
goto plst
:: ==============================================
:pslist
echo PSLIST.txt Found, checking...
!!md5sum %path1%\OUTPUT\PSLIST.txt > $PSLIST.td5
set /p PSLIST= < $PSLIST.td5
set PSLIST=%PSLIST:~1,32%
echo Report: %PSLIST%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr PSLIST.txt > $pslt.td5
set /p pslt= < $pslt.td5
set pslt=%pslt:~1,32%
echo Actual: %pslt%
if %pslt%==%PSLIST% goto pslf
color 0c
set err=1
echo PSLIST.txt Hash Mismatch!!
echo.
pause
goto pf
:pslf
echo Verified!
echo.
goto pf
:: ==============================================
:psfile
echo PSFILE.txt Found, checking...
!!md5sum %path1%\OUTPUT\PSFILE.txt > $PSFILE.td5
set /p PSFILE= < $PSFILE.td5
set PSFILE=%PSFILE:~1,32%
echo Report: %PSFILE%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr PSFILE.txt > $pfle.td5
set /p pfle= < $pfle.td5
set pfle=%pfle:~1,32%
echo Actual: %pfle%
if %pfle%==%PSFILE% goto plef
color 0c
set err=1
echo PSFILE.txt Hash Mismatch!!
echo.
pause
goto si
:plef
echo Verified!
echo.
goto si
:: ==============================================
:sysinfo
echo SYSINFO.txt Found, checking...
!!md5sum %path1%\OUTPUT\SYSINFO.txt > $SYSINFO.td5
set /p SYSINFO= < $SYSINFO.td5
set SYSINFO=%SYSINFO:~1,32%
echo Report: %SYSINFO%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr SYSINFO.txt > $sysi.td5
set /p sysi= < $sysi.td5
set sysi=%sysi:~1,32%
echo Actual: %sysi%
if %sysi%==%SYSINFO% goto sysif
color 0c
set err=1
echo SYSINFO.txt Hash Mismatch!!
echo.
pause
goto pi
:sysif
echo Verified!
echo.
goto pi
:: ==============================================
:psinfo
echo SPSinfo.txt Found, checking...
!!md5sum %path1%\OUTPUT\PSinfo.txt > $PSinfo.td5
set /p PSinfo= < $PSinfo.td5
set PSinfo=%PSinfo:~1,32%
echo Report: %PSinfo%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr PSinfo.txt > $psi.td5
set /p psi= < $psi.td5
set psi=%psi:~1,32%
echo Actual: %psi%
if %psi%==%PSinfo% goto psif
color 0c
set err=1
echo PSinfo.txt Hash Mismatch!!
echo.
pause
goto wi
:psif
echo Verified!
echo.
goto wi
:: ==============================================
:whoami
echo WhoamI.txt Found, checking...
!!md5sum %path1%\OUTPUT\WhoamI.txt > $WhoamI.td5
set /p WhoamI= < $WhoamI.td5
set WhoamI=%WhoamI:~1,32%
echo Report: %WhoamI%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr WhoamI.txt > $who.td5
set /p who= < $who.td5
set who=%who:~1,32%
echo Actual: %who%
if %who%==%WhoamI% goto whof
color 0c
set err=1
echo WhoamI.txt Hash Mismatch!!
echo.
pause
goto mc
:whof
echo Verified!
echo.
goto mc
:: ==============================================
:mac
echo MAC.txt Found, checking...
!!md5sum %path1%\OUTPUT\MAC.txt > $MAC.td5
set /p MAC= < $MAC.td5
set MAC=%MAC:~1,32%
echo Report: %MAC%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr MAC.txt > $mca.td5
set /p mca= < $mca.td5
set mca=%mca:~1,32%
echo Actual: %mca%
if %mca%==%MAC% goto mcaf
color 0c
set err=1
echo MAC.txt Hash Mismatch!!
echo.
pause
goto cl
:mcaf
echo Verified!
echo.
goto cl
:: ==============================================
:curl
echo CURL.txt Found, checking...
!!md5sum %path1%\OUTPUT\CURL.txt > $CURL.td5
set /p CURL= < $CURL.td5
set CURL=%CURL:~1,32%
echo Report: %CURL%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr CURL.txt > $crl.td5
set /p crl= < $crl.td5
set crl=%crl:~1,32%
echo Actual: %crl%
if %crl%==%CURL% goto crlf
color 0c
set err=1
echo CURL.txt Hash Mismatch!!
echo.
pause
goto ed
:crlf
echo Verified!
echo.
goto ed
:: ==============================================
:edd
echo EDD.txt Found, checking...
!!md5sum %path1%\OUTPUT\EDD.txt > $EDD.td5
set /p EDD= < $EDD.td5
set EDD=%EDD:~1,32%
echo Report: %EDD%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr EDD.txt > $eddd.td5
set /p eddd= < $eddd.td5
set eddd=%eddd:~1,32%
echo Actual: %eddd%
if %eddd%==%EDD% goto edddf
color 0c
set err=1
echo EDD.txt Hash Mismatch!!
echo.
pause
goto di
:edddf
echo Verified!
echo.
goto di
:: ==============================================
:driveid
echo DRIVEID.txt Found, checking...
!!md5sum %path1%\OUTPUT\DRIVEID.txt > $DRIVEID.td5
set /p DRIVEID= < $DRIVEID.td5
set DRIVEID=%DRIVEID:~1,32%
echo Report: %DRIVEID%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr DRIVEID.txt > $drvi.td5
set /p drvi= < $drvi.td5
set drvi=%drvi:~1,32%
echo Actual: %drvi%
if %drvi%==%DRIVEID% goto drvif
color 0c
set err=1
echo DRIVEID.txt Hash Mismatch!!
echo.
pause
goto te
:drvif
echo Verified!
echo.
goto te
:: ==============================================
:time_e
echo Time_E.txt Found, checking...
!!md5sum %path1%\OUTPUT\Time_E.txt > $Time_E.td5
set /p Time_E= < $Time_E.td5
set Time_E=%Time_E:~1,32%
echo Report: %Time_E%
type %path1%\OUTPUT\HASH_VALUES.txt |findstr Time_E.txt > $timee.td5
set /p timee= < $timee.td5
set timee=%timee:~1,32%
echo Actual: %timee%
if %timee%==%Time_E% goto timeef
color 0c
set err=1
echo Time_E.txt Hash Mismatch!!
echo.
pause
goto exit
:timeef
echo Verified!
echo.
goto exit


:exit
IF EXIST *.td5 del *.td5
echo.
echo ================
if %err%==1 echo ANOMALY DETECTED
if %err%==0 echo INTEGRITY VERIFIED
echo ================
echo.
echo Check Completed
echo.
pause
color 07