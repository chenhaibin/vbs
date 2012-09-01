@echo off

echo Inserting into registry
regedit /s afinitpassword.reg

echo Changing password
hxmodusr -LOCAL -USER:afbackup -PASSWORD:ueyVfv^Ljk'?iC