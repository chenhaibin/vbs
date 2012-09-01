' AFfoldersize.vbs - Shows how much space is taken by each subdirectory
' Amended by - Adrian Farnell
' Version no - v1.0
' Date Amended - 18th July 2000

on error resume next

' Start the filesystemobject.
set objfso = createobject ("scripting.filesystemobject")

'Get arguements from command-line
set collArgs = wscript.Arguments

strtotalsizeofallfolders=0
strrootfolder = collargs(0)

	wscript.echo ""
	wscript.echo "	USAGE: cscript AFfoldersize.vbs [root_folder]"
	Wscript.echo ""
	Wscript.echo "		[Root_folder] is the folder with"
	wscript.echo "		which to start enumeration from"
	wscript.echo ""

if strrootfolder="" then 
	wscript.echo  "No Arguments specified"
	wscript.quit
end if

'Get selected root folder
set objfolder = objfso.getfolder (strrootfolder)


wscript.echo ""
wscript.echo "Starting from folder: " & strrootfolder
wscript.echo ""

'make a folder collection
set collfolder = objfolder.subfolders

'Display Table
wscript.echo "Folder				Size/Mb"
wscript.echo "-----------------------------------------"


'Enumerate each folder, displaying its name and size

For each folder in collfolder

	strlengthofname=len(folder.name)
	if strlengthofname<8 then strtabs=chr(9) & chr(9) & chr(9) 
	if strlengthofname>7.5 and strlengthofname<15.5 then strtabs=chr(9) & chr(9)
	if strlengthofname>15.5 and strlengthofname<23.5 then strtabs=chr(9)
	if strlengthofname>23.5 then strtabs=chr(32) 'else strtabs=chr(32)
  	
	strformattedsize = folder.size / 1048576
	strformattedsize = formatnumber (strformattedsize, 3)
	Wscript.echo folder.name & strtabs & chr(9) & strformattedsize & chr(9)
	strtotalsizeofallfolders = strtotalsizeofallfolders + strformattedsize

Next

wscript.echo ""
wscript.echo "Total Size/Mb: " & strtotalsizeofallfolders