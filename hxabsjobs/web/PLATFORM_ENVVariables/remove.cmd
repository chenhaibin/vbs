@echo off

:: *****************************************************
:: Name		: Remove.cmd 
:: Author	: Adrian Farnell
:: Version	: 1.0
:: Notes	: removes the environment variable, application
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
cscript %1\hxremovevariables.vbs

goto end


:errusage
Echo.
Echo  Usage: remove.cmd [install_dir]
Echo.
goto end

:end