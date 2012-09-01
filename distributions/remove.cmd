@echo off
:: remove.cmd - removes all traces of winvnc
:: version: 1.0
:: author : Adrian Farnell
:: date : 15.01.2000

::***********************************************
:: stop winvnc service
::***********************************************

echo  Stopping WinVNC service
net stop winvnc

::***********************************************
:: del registry entries
::***********************************************

echo  Deleting Registry Entries


::HKLM\Software\ORL
::HKU\Software\ORL
::HKLM\System\CurrentControlSet\services\winvnc


::***********************************************
:: del files from e:\halifax\utils
::***********************************************

echo  Deleting files directory
del /s e:\halifax\utils\winvnc


:end