:: File Name: HxRegDump.cmd
:: Version: 1.0
:: Date Amended: 21.09.2000
:: Who Amended: Adrian Farnell

::*****************************************************
:: Dumps old registry key for use if we need to backout
:: Called by hxipreaddresser.vbs
::*****************************************************

regdmp.exe HKEY_LOCAL_MACHINE\System\Controlset001\Services\%1\Parameters\Tcpip > %computername%.txt
