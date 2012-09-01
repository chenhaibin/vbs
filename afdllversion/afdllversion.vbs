' afdllversion.vbs - Displays version information for a selected dll(s)
' Version 	: pre-alpha
' Date		: 04.01.2000

Option Explicit

on error resume next

dim objfso, objdllsini, objwsharguments, objversioninfo, xobj
dim strUSAGE, strNAMEofDLL, blRESULT, strout, akeys, i, strvalue, strResult, strcompanyname
dim strfiledescription, strfileversion, strinternalname, strlegalcopyright, stroriginalfilename, strproductname
dim strproductversion, strfilename, strmajorversion, strminorversion, strfileflags, strfileos, strfiletype
dim strline, cresult, strsizeoffile, strdatestamponfile,strtabs

'start filesystemobject
set objfso = wscript.createobject ("scripting.filesystemobject")

'start versioninfo object
set objversioninfo = wscript.createobject ("softwing.versioninfo")

'start arguments
set objwsharguments = wscript.arguments

'Usage
strUSAGE = " afdllversion.vbs - Script to determine dll versions" & vbcrlf & _
		     "                    Can be used in conjunction with a dlls.ini" & vbcrlf & _
		     "                    or just a specific dll" & vbcrlf & vbcrlf & _
		     " [USAGE] - cscript afdllversion.vbs [name of dll] /useini" & vbcrlf & vbcrlf & _
		     " [name of dll] optional. Specifies a particular dll to examine." & vbcrlf & _
		     "               If left, script uses dlls.ini."
	


if objwsharguments.count <> 1  then
	
	wscript.echo strUSAGE
	wscript.quit
			
end if

if objwsharguments.item(0) = "/?" then
	
	wscript.echo strUSAGE
	wscript.quit
			
end if

strNAMEofDLL = objwsharguments.item(0)



if objwsharguments.item(0) = "/useini" then
	
	wscript.echo " FileName" & chr(9) & chr(9) & chr(9) & "Prod Ver." & chr(9) & "File Ver." & chr(9) & "Maj.Min ver" & chr(9) & "Date" & chr(9) & "Size"
	wscript.echo " ------------------------------------------------------------------------------------------------"
	
	'Open ini file as a file
	set objdllsini = objfso.opentextfile ("dlls.ini", 1)

	if err.number <> 0 then
		wscript.echo " Error! Can not find dlls.ini"
		wscript.quit
	end if

	'Multi dll query
	do while objdllsini.atendofstream <> true
		
		strline = objdllsini.readline
		'wscript.echo objdllsini.readline
		if Left (strline, 1) = "#" then
			'wscript.echo "ignoring"
			if trim(objdllsini.readline) = "" then 
			'wscript.echo "blankline"
			end if
		else
			'wscript.echo "n"
			
			getfiledetails trim(strline)
			if Len(strproductversion) < 4 then strtabs = chr(9) & chr(9) else strtabs=chr(9)
			wscript.echo " " & strfilename & chr(9) & strproductversion & strtabs & strfileversion & chr(9) & strMajorversion & "." & strminorversion & chr(9) & chr(9) & strdatestamponfile & chr(9) & strsizeoffile
		end if 
	loop	
	wscript.quit
else
	'Single dll query
	getfiledetails strNAMEofDLL
	wscript.echo " Company Name     : " & strcompanyname
	wscript.echo " File Description : " & strfiledescription
	wscript.echo " File Version     : " & strfileversion
	wscript.echo " Internal Name    : " & strinternalname
	wscript.echo " Legal Copyright  : " & strlegalcopyright
	wscript.echo " Original Filename: " & stroriginalfilename
	wscript.echo " Product Name     : " & strproductname
	wscript.echo " Product Version  : " & strproductversion
	wscript.echo " File Name        : " & strfilename
	wscript.echo " Major Version    : " & strmajorversion
	wscript.echo " Minor Version    : " & strminorversion
	wscript.echo " File Flags       : " & strfileflags
	wscript.echo " File OS          : " & strfileos
	wscript.echo " File Type        : " & strfiletype
	wscript.quit
end if




'--------------------------------------------------------------------------------------
function getfiledetails(strother)

	blresult = objversioninfo.getbyfilename(strother)
	
	

	
		if blresult <> 1 then
			wscript.echo " Error! File " & strother & " Can not be loaded."
		
		else
			strcompanyname = objversioninfo.getvalue("CompanyName")
			strfiledescription = objversioninfo.getvalue("FileDescription")
			strfileversion = objversioninfo.getvalue("FileVersion")
			strinternalname = objversioninfo.getvalue("InternalName")	
			strlegalcopyright = objversioninfo.getvalue("LegalCopyright")
			stroriginalfilename = objversioninfo.getvalue("OriginalFileName")
			strproductname = objversioninfo.getvalue("ProductName")
			strproductversion = objversioninfo.getvalue("ProductVersion")
			
			strfilename = objversioninfo.Filename
			strmajorversion = objversioninfo.Majorversion
			strminorversion = objversioninfo.MinorVersion
			strfileflags = objversioninfo.FileFlags
			strfileos = objversioninfo.FileOS
			strfiletype = objversioninfo.FileType


		end if
	
	set cresult = objfso.getfile(strother)
	'wscript.echo strother
	'if cresult <> 1 then 
	'	wscript.echo " Error! File " & strother & " Can not be loaded."
	'else
		strsizeoffile = cresult.size
		strdatestamponfile = cresult.datecreated
		'wscript.echo strsizeoffile
		'wscript.echo strdatestamponfile
	'end if
end function
'-----------------------------------------------------------------------------------------



