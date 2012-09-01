' FILE: createweb.vbs
' DESC: This script illustrates how you can programmatically
'       create a web site using IIS administration objects.
' AUTH: Thomas L. Fredell
' DATE: 11/23/98
'
' Copyright 1998 Macmillan Publishing

Option Explicit

On Error Resume Next

Dim oArgs, sComputer, sHostName, sRootDir
Dim oService, nNewServer, oServer, oVDir, oDir
Dim oFS, sBasePath, sCGIDir
Dim sTmp

' Step 1: Check script arguments
Set oArgs = Wscript.Arguments
If oArgs.Count <> 4 Then
	Wscript.Echo "USAGE:  createweb  [computer] [hostname]" & _
					" [instance] [rootdir]"
	Wscript.Echo ""
	Wscript.Echo "        Creates a new IIS web site on "
	Wscript.Echo "        [computer] using the [hostname]."
	Wscript.Echo "        [instance] specifies the web server"
	Wscript.Echo "        and [rootdir] specifies the root"
	Wscript.Echo "        directory on the file system."
	Wscript.Echo ""
	Wscript.Quit 1
End If
sComputer = oArgs(0)
sHostName = oArgs(1)
nNewServer = oArgs(2)
sRootDir = oArgs(3)

' Step 2: Get the object that represents the web service
Set oService = GetObject("IIS://" & sComputer & "/W3SVC")

' Step 3: Create a new virtual server
' Determine the new virtual server number by
' iterating through the current servers
oService.GetInfo
Set oServer = oService.Create("IIsWebServer", nNewServer)
If Err.Number <> 0 Then
	ShowErr "Unable to create server."
	Wscript.Quit Err.Number
End If

' Step 4: Configure the new server
oServer.DefaultDoc = "default.htm, index.htm"
oServer.ServerComment = "New server created by createweb.vbs"
oServer.ConnectionTimeout = 600
oServer.ServerBindings = ":80:" & sHostName
oServer.SetInfo
If Err.Number <> 0 Then
	ShowErr "An error occurred while setting configuration params."
	Wscript.Quit Err.Number
End If

' Step 5: Create virtual root directory for the new site
Set oVDir = oServer.Create("IIsWebVirtualDir", "ROOT")
If Err.Number <> 0 Then
	ShowErr "Unable to create virtual root directory."
	Wscript.Quit Err.Number
End If

' Step 6: Configure the virtual root directory
Set oFS = CreateObject("Scripting.FileSystemObject")
sBasePath = oFS.GetAbsolutePathName(sRootDir)
oVdir.Path = sBasePath
oVdir.AccessRead = True
oVdir.AccessWrite = True
oVdir.AccessScript = True
oVdir.SetInfo
If Err.Number <> 0 Then
	ShowErr "Unable to save settings for virtual root."
	Wscript.Quit Err.Number
End If
oVdir.AppCreate False
If Err.Number <> 0 Then
	ShowErr "Unable to define virtual root as application."
	Wscript.Quit Err.Number
End If
oVdir.SetInfo

' Step 7: Create and configure a cgi-bin directory
' First ensure that the physical directory exists
sCGIDir = sBasePath & "\cgi-bin"
If Not oFS.FolderExists(sCGIDir) Then
	' Create the physical cgi-bin directory
	oFS.CreateFolder sCGIDir
End If

' Then create the IIsWebDirectory object & configure it
Set oDir = oVDir.Create("IIsWebDirectory", "cgi-bin")
If Err.Number <> 0 Then
	ShowErr "Unable to create cgi-bin directory."
	Wscript.Quit Err.Number
End If
oDir.AccessRead = False
oDir.AccessWrite = False
oDir.AccessScript = True
oDir.AccessExecute = True
oDir.SetInfo
If Err.Number <> 0 Then
	ShowErr "Error setting cgi-bin configuration."
	Wscript.Quit Err.Number
End If

' Step 8: Start the web server
oServer.Start
If Err.Number <> 0 Then
	ShowErr "Unable to start new web server."
	Wscript.Quit Err.Number
End If

' Step 9: Done! Quit execution
Wscript.Echo "createweb.vbs: Successfully created new web site."
Wscript.Echo "createweb.vbs: Instance #=" & nNewServer
Wscript.Echo "createweb.vbs: Host Name=" & sHostName
Wscript.Echo "createweb.vbs: Root Dir=" & sRootDir
Wscript.Quit 0

'------------------------------------------------------
' SUB:  ShowErr(Description)
' DESC: Displays information about a runtime error.
'------------------------------------------------------
Sub ShowErr(sDescription)
	Wscript.Echo "createweb.vbs: " & sDescription
	Wscript.Echo "createweb.vbs: An error occurred - code: " & _
		Hex(Err.Number) & " - " & Err.Description
End Sub
