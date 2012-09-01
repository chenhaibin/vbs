
' create lots of little files

strNoFiles = 10
strMin = 97
strMax = 122
strDestination = ".\Lots\"

set oFSO = createobject ("scripting.filesystemobject")

for a = 1 to 105 'makes about 1MB
	subCreateFile
next


sub subCreateFile

	For i = 1 to 8 

		Randomize 
		
		StrChr = ((strMax - strMin) * Rnd() + strMin)
		
		strName = strName & Chr(strChr)

	Next 

	if ofso.fileexists ( strDestination & strName & ".txt") then 
		subCreateFile
	else
		ofso.copyfile "copy.tmp", strDestination & strName & ".txt"
	end if
	

end sub

set oFSo = nothing