@echo off
:: Name		: HxRestoreIPs.cmd - Restores old IP address settings to NIC (after Hxipreaddresser.vbs)
:: VErsion	: pre-alpha
:: Date amended	: 14.09.2000
:: Last amended by: Adrian Farnell

set NEWUSER=%1

if "%1"=="" goto DISPLAYUSAGE
if not exist %computername%.txt goto FILEERROR

::***************************************************
::Run regini.exe to restore old registry NIC settings
::***************************************************

regini.exe %computername%.old

::***************************************************
:: Display registry settins which have just changed
::***************************************************

regdmp.exe HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\%1\Parameters\TcpIp
echo. | Date >>log.txt
echo. | Time >> log.txt
echo Restored old IP settings >>log.txt
regdmp.exe HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\%1\Parameters\TcpIp >> log.txt

echo.
echo ************************************************
echo Please Reboot server to activate changes
echo ************************************************
echo.

goto end

::****************************************************
:: Usage
::****************************************************

:DISPLAYUSAGE

echo.
echo  HxRestoreIPs.cmd - script to restore previous NIC settings
echo.
echo 	USAGE: 	HxRestoreIPs.cmd [NIC]
echo.
echo  Where NIC is the registry value for the NIC
echo.
goto END

::*****************************************************
:: File doesn't exist error
::*****************************************************

:FILEERROR
echo. 
echo %computername%.old does not exist. HxIPReaddresser.vbs has 
echo never been run.
echo. 
echo Please contact ECP for help
goto end

:END