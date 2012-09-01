
set objReg = createobject ("a1asp.reg")

objReg.Rootkey = HKEY_LOCAL_MACHINE
objReg.OpenKey "\SYSTEM\CurrentControlSet\Enum\PCI", false

n = objReg.GetKeyNames
wscript.echo n

if n > 0 then 
	
	for i = 0 to n - 1
		wscript.echo objReg.GetKeyname(i)
	next
end if

objReg.CloseKey

set objReg = nothing