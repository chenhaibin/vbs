::echo. >> md5s.txt
::echo %1 >> md5s.txt
::echo ---------- >> md5s.txt
::echo. >> md5s.txt

md5sum "%1"/* >> md5s.txt