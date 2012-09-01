:: Name of File: %computername%.cmd
:: Version: 1.0
:: Date Amended: 21.09.2000
:: Who Amended: Adrian Farnell

:: ********************************************
:: This File uses the command hxipreaddresser
:: to assign the correct ipaddresses for the name
:: of the server
:: ********************************************

hxipreaddresser.vbs /start: /finish: /gateway: /Subnet: /NIC

:: ********************************************
:: Update history.txt
:: ********************************************

cscript //nologo e:\halifax\scripts\hist_log.vbs "Changing IP address on server via hxipreaddresser.vbs"