@echo off

::***************************************************
:: Filename 	: HxAccessPerfmon.cmd
:: 		  (Allows read access to perfmon stats
::                to specified users).
::
:: Date		: 09.10.2000
:: Author	: Adrian Farnell
:: Version	: v1.0
::
:: Dependencies : Needs Secadd.exe and xcacls.exe
:: 		  from NT resource kit to work
::***************************************************

set NEWUSER=%1

if "%1"=="" goto DISPLAYUSAGE
if not exist c:\winnt\system32\perfc009.dat goto FILENOTTHERE
if not exist c:\winnt\system32\perfh009.dat goto FILENOTTHERE
if "%1"=="/DISPLAY" goto CURRENTSETTINGS
if "%2"=="/DISPLAY" goto CURRENTSETTINGS
if "%1"=="/display" goto CURRENTSETTINGS
if "%2"=="/display" goto CURRENTSETTINGS
if "%1"=="/?" goto DISPLAYUSAGE

:: ***************************************************
:: log the current details of the dat files to a log
:: ***************************************************

echo ************* >> logfile.txt
echo. >> logfile.txt
echo. | date >> logfile.txt
echo. | time >> logfile.txt
echo. >> logfile.txt
xcacls c:\winnt\system32\perfc009.dat >> logfile.txt
xcacls c:\winnt\system32\perfh009.dat >> logfile.txt

:: ***************************************************
:: give access to the specified user to the .dat files
:: ***************************************************

xcacls c:\winnt\system32\PERFC009.dat /E /G %NEWUSER%:R >> logfile.txt

xcacls c:\winnt\system32\PERFH009.dat /E /G %NEWUSER%:R >> logfile.txt

:: ***************************************************
:: Check file permissions after change
:: ***************************************************
echo. >>logfile.txt

xcacls c:\winnt\system32\perfc009.dat >> logfile.txt
xcacls c:\winnt\system32\perfh009.dat >> logfile.txt

:: ***************************************************
:: give read access to world users using regini.exe
:: ***************************************************

secadd -l -a "Software\Microsoft\Windows NT\CurrentVersion\Perflib" %NEWUSER% >> logfile.txt
secadd -l -a "System\CurrentControlSet\Control\SecurePipeServers\winreg" %NEWUSER% >> logfile.txt

goto END

:DISPLAYUSAGE
echo.
echo  HxAccessPerfmon.cmd - script to allow users access to perfmon stats
echo.
echo 	USAGE: 	HxAccessPerfmon.cmd [Domain\User_to_add] /DISPLAY
echo.
echo    [Domain\User_to_Add] - Used to add specific user to 
echo                           perfmon access list
echo    /DISPLAY             - Used to display current list of users
echo                           who have access. Should not be used at 
echo                           the same time as adding a user
goto END

:FILENOTTHERE
echo. 
echo Some Files may not be present, script aborted
echo Some Files may not be present, script aborted >> logfile.txt
goto END

:CURRENTSETTINGS

:: ****************************************************
:: display the current permissions of files
:: ****************************************************

echo.
echo Displaying current details:
echo. 
xcacls c:\winnt\system32\perfc009.dat
xcacls c:\winnt\system32\perfh009.dat

goto end

:: *****************************************************
:: display current settings on registry keys
:: *****************************************************

:END