@echo off
:: after.cmd - Installs VNC as a service, does not alter any dlls
:: author: Adrian Farnell
:: date: 15.01.2000
:: version: 1.0

set source=%1

::************************************************
:: If source is blank display usage
::************************************************

if "%1"=="" goto noargument

::************************************************
:: Copies install files to another location
::************************************************

echo  Source Directory is: %source%

::*************************************************
:: Check for directory structure
::*************************************************

if not exist e: goto noedrive

if not exist e:\halifax echo  ** creating e:\halifax

if not exist e:\halifax md e:\halifax

if not exist e:\halifax\utils echo  ** creating e:\halifax\utils

if not exist e:\halifax\utils md e:\halifax\utils

if not exist e:\halifax\utils\winvnc echo  ** creating e:\halifax\utils\winvnc

if not exist e:\halifax\utils\winvnc md e:\halifax\utils\winvnc

::**************************************************
:: Copy files over
::**************************************************

date /t >> Install.log
time /t >> install.log
copy %source%\*.* e:\halifax\utils\winvnc >>install.log

::*************************************************
:: Passwords into registry
::*************************************************

echo  Entering password into registry
regedit /s e:\halifax\utils\winvnc\file.reg
regedit /s e:\halifax\utils\winvnc\file2.reg

::*************************************************
:: Start winvnc install
::*************************************************

echo  Install WinVNC as a service
e:\halifax\utils\winvnc\winvnc.exe -install

echo  Starting WinVNC service
net start winvnc >> install.log

goto end

:noargument
::************************************************
:: Error for no argument
::************************************************
echo.
echo Error! Source directory not specified
goto displayusage

:noedrive
::************************************************
::Error for no E: drive
::************************************************
echo. Error! No E drive
goto End

::************************************************
:: Displays usage if there is an error
::************************************************

:displayusage
echo.
echo  USAGE: after.cmd [source_dir]
echo.
echo  Where source_dir is where the source files for install are kept 
goto end

:end
time /t >> install.log