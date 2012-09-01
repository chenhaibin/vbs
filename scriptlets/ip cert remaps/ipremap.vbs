'****************************************************************************
'
'  SCRIPT FILE
'   IPAddressMapping
'
'  SYNOPSIS
'   This script reads Bobbafat's IP Address mapping file and returns
'	the IP Address required in the form of a batch file which when
'	executed, sets an environment variable.
'
'  PARAMETERS
'   strServerName - the name of the web server, eg 'ECOM1201'
'   strWebName    - the description of the web site, eg 'WWWUSERREG"
'
'  NOTE
'   The caller does not need to specify the 'L1' or 'L2' which is how the
'   web names appear in the csv file. However, if L1 or L2 is specified
'	on the end of the site name, the function will still function fine.
'
'  AUTHOR
'   Steve Mitchell (1.3 additions Chris Day
'
'  CHANGE HISTORY
'   25th August         1999    -   (1.0) Creation
'   27th August         1999    -   (1.1) Amended to not be case-sensitive, and
'                                         also to remove any unnecessary zeros
'                                         from the IP Address (SRM)
'   31st August         1999    -   (1.2) Added some fairly cunning error reporting
'                                         in the form of adding @ECHO commands to the
'                                         temp batch file. Cock-a-doodle-do (SRM)
'   15th August         2000    -   (1.3) Added support for user specified ipmapfile
'
'****************************************************************************
Option Explicit

Dim oArgs
Dim ErrNum
Dim strIPAddress
Set oArgs = WScript.Arguments

If oArgs.count < 3 then
	WScript.Echo "Usage: cscript IPAddressMapping.vbs ServerName WebName IPMapFile"
	WScript.Echo ""
	WScript.Echo "Where,"
	WScript.Echo " ServerName - name of the physical server"
	WScript.Echo " WebName    - description of the web site"
	WScript.Echo " IPMapFile - name of the ipmapfile to use"
	WScript.Echo ""
	WScript.Echo "Example,"
	WScript.Echo "  cscript IPAddressMapping.vbs ECOM1201 WWWUSERREG IPMapFile"
	WScript.Quit 1
End If

'WScript.Echo "Searching IP mapping file for Server=" & oArgs(0) & ", Web=" & oArgs(1) & "..."
strIPAddress = GetIPAddress(oArgs(0), oArgs(1), oArgs(2))
'WScript.Echo "  IP Address = " & strIPAddress
WScript.echo "@SET TMP_SITE_IP=" & strIPAddress

'****************************************************************************
' Check for runtime errors
'****************************************************************************
If Err.Number <> 0 Then
	ErrNum = Err.Number
	WScript.Echo "@ECHO ERROR!!! Installation of package " & PakName & " failed!"
	WScript.Echo "@ECHO Number=" & Err.Number & ", Description=" & Err.Description & ", Source=" & Err.Source
	Err.Clear
	WScript.Quit Abs(ErrNum)
End If

'****************************************************************************
'
'  FUNCTION
'   GetIPAddress(strServerName, strWebName, IPMapFile)
'
'  SYNOPSIS
'   This function reads in the IP Address file and searches for the
'   details specified
'
'  PARAMETERS
'   strServerName - the name of the web server, eg 'ECOM1201'
'   strWebName    - the description of the web site, eg 'WWWUSERREG"
'   IPMapFile     - the ipmapfile to use, eg 'ecomipmap.txt"
'
'  AUTHOR
'   Steve Mitchell
'
'  CHANGE HISTORY
'   25th August         1999    -   (1.0) Creation
'   15th August         2000    -   (1.3) Added support for user specified ipmapfile
'
'****************************************************************************
Function GetIPAddress(strServerName, strWebName, IPMapFile)
    On Error Resume Next
	Dim ErrNum
	Dim sLineInput
    Dim bolFound
    Dim strIPAddress
    Dim strSearch
	Dim objFileSystem
	Dim objFile
	Dim strFileName
	Dim intPos
    
    strSearch = LCase(strServerName & "," & strWebName)
    bolFound = False
    strFileName = "C:\WINNT\SYSTEM32\DRIVERS\ETC\" &IPMapFile
    'WScript.Echo strFileName

    '*************************************************************************
    ' Create the FileSystem object and ensure that the mapping file exists
    '*************************************************************************
	Set objFileSystem = CreateObject("Scripting.FileSystemObject")
	If Err.Number <> 0 Then
		ErrNum = Err.Number
		WScript.Echo "@ECHO   ERROR!!! Cannot create FileSystemObject"
		WScript.Echo "@ECHO   Number=" & Err.Number & ", Description=" & Err.Description & ", Source=" & Err.Source
		Err.Clear
		WScript.Quit ErrNum
	End If
	If objFileSystem.FileExists(strFileName) = False Then
		WScript.Echo "@ECHO   ERROR!!! IP Address Mapping file does not exist in C:\WINNT\SYSTEM32\DRIVERS\ETC"
		Set objFileSystem = Nothing
		WScript.Quit 1
	End If
    
    '*************************************************************************
    ' Open the file
    '*************************************************************************
'	WScript.Echo "  Opening mapping file..."
    Set objFile = objFileSystem.OpenTextFile (strFileName, 1, False)
	If Err.Number <> 0 Then
		ErrNum = Err.Number
		WScript.Echo "@ECHO     ERROR!!! Cannot open the file, although the file DOES exist."
		WScript.Echo "@ECHO     Number=" & Err.Number & ", Description=" & Err.Description & ", Source=" & Err.Source
		Err.Clear
		Set objFileSystem = Nothing
		WScript.Quit Abs(ErrNum)
	End If

    '*************************************************************************
    ' Check line be line til we find our server. Detect EOF be trapping an
    ' error on the ReadLine method.
    '*************************************************************************
    Err.Clear
    Do	' Sorry Rich, but it had to be done :)
        sLineInput = LCase(objFile.ReadLine)
        If Err.Number <> 0 Then
			Err.Clear
			Exit Do
        End If

        If sLineInput <> "" Then
            If InStr(1, sLineInput, strSearch) > 0 Then
				intPos = InStrRev(sLineInput, ",") + 1
                strIPAddress = Mid(sLineInput, intPos, Len(sLineInput) - intPos + 1)
                bolFound = True
                Exit Do
            End If
        End If
    Loop
	objFile.Close
	Set objFileSystem = Nothing
    
    '*************************************************************************
    ' See if we've found an IP Address
    '*************************************************************************
    If bolFound Then
        GetIPAddress = TidyIPAddress(strIPAddress)
    Else
'        GetIPAddress = "0.0.0.0"
		WScript.Echo "@ECHO     ERROR!!! IP Address cannot be found in the file."
		WScript.Quit 1
    End If
    
    '*************************************************************************
    ' If a runtime error occurred, send a blank IP address
    '*************************************************************************
    If Err.Number <> 0 Then
		WScript.Echo "@ECHO     ERROR!!! IP Address cannot be found in the file."
		WScript.Quit 1
'        GetIPAddress = "0.0.0.0"
    End If
End Function

'****************************************************************************
'
'  FUNCTION
'   TidyIPAddress(strIPAddress)
'
'  SYNOPSIS
'   This function removes any unnecessary zeros as these can shag
'   IIS's ability to understand the IP Address. D'oh!
'
'  PARAMETERS
'   strIPAddress  - exactly what it says on the tin
'
'  AUTHOR
'   Steve Mitchell
'
'  CHANGE HISTORY
'   27th August         1999    -   (1.0) Creation
'
'****************************************************************************
Function TidyIPAddress(strIPAddress)
	On Error Resume Next
	Dim i
	'*************************************************************************
	' Remove any zeros that follow dots.
	' Also, remove any leading zeros from the forst number
	'*************************************************************************
    For i = 1 To 3
        strIPAddress = Replace(strIPAddress, ".0", ".")
	If Left(strIPAddress, 1) = "0" Then strIPAddress = Right(strIPAddress, Len(strIPAddress) - 1)
    Next

	'*************************************************************************
	' Now check that we haven'tgone too far. If we've removed any genuine zeros,
	' put them back in.
	'*************************************************************************
    strIPAddress = Replace(strIPAddress, "..", ".0.")
    strIPAddress = Replace(strIPAddress, "..", ".0.")  'This needs doing twice!!
    If Left(strIPAddress, 1) = "." Then strIPAddress = "0" & strIPAddress
    If Right(strIPAddress, 1) = "." Then strIPAddress = strIPAddress & "0"
    TidyIPAddress = strIPAddress
End Function
