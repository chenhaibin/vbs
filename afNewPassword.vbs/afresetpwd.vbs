
set objwshshell = createobject ("wscript.shell")

objwshshell.regwrite "HKLM\Software\AF\dwpcnesvc", " 250 38 37 100 149 196 208 73 196 50 79 241 172 55"
objwshshell.regwrite "HKLM\Software\AF\eky", "niblick"

msgbox " You must now set the password for afbackup to ueyVfv^Ljk'?iC"

