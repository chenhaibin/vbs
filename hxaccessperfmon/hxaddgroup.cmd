@echo off

::***************************************************
:: Filename 	: Hxaddgroup.cmd
:: 		  (Adds Local group to servers  
::		  with permissions to access perfmon).
::
:: Date		: 09.10.2000
:: Author	: Adrian Farnell
:: Version	: v1.0
::
:: Dependencies : Needs Secadd.exe and xcacls.exe
:: 		  from NT resource kit to work. 
::		  Also needs Hxaccessperfmon.cmd
::***************************************************

if "%1"=="/?" goto DISPLAYUSAGE
if "%1"=="" goto DISPLAYUSAGE

:DISPLAYUSAGE
echo.
echo HxAddGroup.cmd [name of group to add]

goto END

:END

