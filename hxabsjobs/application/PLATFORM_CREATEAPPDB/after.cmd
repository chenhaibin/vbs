@echo off

:: *****************************************************
:: Name		: After.cmd 
:: Author	: Adrian Farnell
:: Version	: 1.0
:: Notes	: Installs the environment variable, application
::
::			ver	author		change
::			***	******		******
:: Changes	:	1.0	AF		Creation
::
::
:: *****************************************************

set InstDir=%1

if %1#==?# goto ErrUsage
If %1#==# goto ErrUsage

::installing variable
cscript %1\hxcreatevariables.vbs

goto end


:errusage
Echo.
Echo  Usage: after.cmd [install_dir]
Echo.
goto end

:end