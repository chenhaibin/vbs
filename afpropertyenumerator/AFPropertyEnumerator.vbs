Const ForReading = 1, ForWriting = 2, ForAppending = 8

' bind x to the object you wish to find optional values for
Set x = GetObject("WinNT://10.208.23.223/administrator")



set fso = CreateObject("Scripting.FileSystemObject")
set Response = fso.OpenTextFile("d:\temp\response.txt", ForWriting, True)

'Get more information on object schema information

Set cls = GetObject(x.Schema)
'Response.Write "Class Name is: " & cls.Name
wscript.echo "Class Name is: " & cls.Name 

For Each op In cls.OptionalProperties

	'Response.Write vbcrlf &"Optional Property: " & op 
	wscript.echo "Optional Property: " & op 

Next

for Each op In cls.MandatoryProperties

	wscript.echo "Mandatory Property: " & op

Next