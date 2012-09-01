
' Name		: chkSecurity.vbs - Checks the guest account is disabled. 
' Author	: Adrian Farnell
' Date		: 03.09.2002
' Version	: 1.0

Option Explicit

Dim oUser, oWshnetwork
Dim strLocal
Dim boolAction
Dim boolDebug
Dim nArgs
Dim boolShare, boolAnyShares

boolShare = FALSE


' collect arguments from command line

if wscript.arguments.count > 3 then 
	subMsg " Too many Arguments"
	subDisplayUsage	
	wscript.quit
else 
	'carry on.....from arguments decide what to do.
	for nArgs = 0 to wscript.arguments.count - 1
		select case wscript.arguments(nArgs)
		
			case "/?"
				subDisplayUsage
			case "/action"
				boolAction = TRUE
			case "/debug"
				boolDebug = TRUE
			case else
				boolAction = FALSE
				boolDebug = FALSE
		end select
	next
end if 

' ** get computername **
set oWshNetwork = createobject ("wscript.network")
strLocal = oWshNetwork.computername

' ** display some debug output **
if boolAction = TRUE then
	subDebugMsg "Script performing fixes"
else
	subDebugMsg "Script will not perform fixes"
end if

' ** display some debug output **
if boolDebug = TRUE then
	subDebugMsg "Script displaying debug text"
else
	subDebugMsg "Script will not display debug text"
end if

' ** Check if guest is disabled **
' ** if not, disable it         **
'subCheckGuest strLocal

' ** Check Users ** 
'subCheckNoUsers strLocal

' ** Check Screen Saver requires password to unlock **
' ** need to check both user and default user **
'subCheckScreenSaverLockout

' **                 Does Administrator exist                       **
' **       Check Administrator is in Administrators Group           **
' ** If administrator exists, and is in the admin group, rename him **
' **              also create dummy admin user                      **
'subCheckAdministrator strLocal

' ** Are there any shares on the machine            **
' ** if so disable admin shares by adding reg entry **
' ** and remove all extraneous shares               **
'subCheckShares strLocal

' ** Check password policy **
' ** If password policy is to weak, will strengthen **
'subCheckPwdPolicy 

' ** Check Partitions are NTFS **
'subCheckNTFS

' ** Check for Modems ** 
'subCheckModems

' ** Check Audit Policy **
'subCheckAuditPolicy

' ** Check for direct writes to hardware ** 
'subCheckDirectDraw

' ** Check if a dumpfile gets generated on an error **
'subCheckDumpFiles

' ** Check temp directories are encrypted **
'subchecktempforencryption

' ** Close script gracefully **
set oWshNetwork = nothing
wscript.quit(0)

' ************ subs and funcs **************

sub subCheckPmonElogACLs

end sub

sub subChangePmonElogACLs

end sub

sub subCheckDirectDraw

	Dim oWshShell
	
	Dim strDirectX

	subMSG "TEST, Direct writes to hardware Memory"

	set oWshShell = createObject ("wscript.shell")
	
	on error resume next
	
	err.clear
	
	strDirectX = oWshShell.regRead ("HKLM\SYSTEM\CurrentControlSet\Control\GraphicsDrivers\DCI\Timeout")
	
	if err.number <> 0 then
		subErrMsg "Registry Entry could not be read"
	end if 
	
	on error goto 0
	
	if strDirectX <> 0 then
		subMSG "WARNING, DirectDraw is enabled"
		subChangeDirectDraw
	else
		subMSG "SECURE, DirectDraw has been disabled"
	end if
	
	set oWshShell = nothing
	

end sub

sub subChangeDirectDraw

	Dim oWshShell

	if boolAction = FALSE then

		'do nothing
	else
	
		subMsg "ACTION, Changing Access to DirectDraw"	
	
		set oWshShell = createObject ("wscript.shell")
		
		on error resume next

		err.clear

		oWshShell.regWrite "HKLM\SYSTEM\CurrentControlSet\Control\GraphicsDrivers\DCI\Timeout", 0, "REG_DWORD"

		if err.number <> 0 then
			subErrMsg "Registry Entry could not be written"
		end if 

		on error goto 0

		subCheckDirectDraw		
		
		set oWshShell = nothing
	
	end if

end sub

sub subCheckDumpFiles

	Dim oWshShell
		
	Dim strCrashDump

	set oWshShell = createObject ("wscript.shell")
	
	subMSG "TEST, Do dump files get created"

	set oWshShell = createObject ("wscript.shell")
	
	on error resume next
	
	err.clear
	
	strCrashDump = oWshShell.regRead ("HKLM\SYSTEM\CurrentControlSet\Control\CrashControl\CrashDumpEnabled")
	
	if err.number <> 0 then
		subErrMsg "Registry Entry could not be read"
	end if 
	
	on error goto 0
	
	Select Case strCrashDump 
	
		case 0
			subMsg "SECURE, No Crash dumps are generated"
		case 1
			subMsg "WARNING, Complete Memory Dump Enabled"
			subChangeDumpFile
		case 2
			subMsg "WARNING, Kernel Dump Enabled"
			subChangeDumpFile
		case 3
			subMsg "WARNING, Mini Memory Dump Enabled"
			subChangeDumpFile
		case else
			subMsg "UNKNOWN, Crash dump value unknown"
			subChangeDumpFile
		
	end select
		
	set oWshShell = nothing

end sub

sub subChangeDumpFile

	Dim oWshShell

	if boolAction = FALSE then

		'do nothing
	else
	
		subMsg "ACTION, Changing Dumpfile creation properties"	
	
		set oWshShell = createObject ("wscript.shell")
		
		on error resume next

		err.clear

		oWshShell.regWrite "HKLM\SYSTEM\CurrentControlSet\Control\CrashControl\CrashDumpEnabled", 0, "REG_DWORD"

		if err.number <> 0 then
			subErrMsg "Registry Entry could not be written"
		end if 

		on error goto 0

		subCheckDumpFiles		
		
		set oWshShell = nothing
	
	end if

end sub

sub subCheckPageFileatShutDown

end sub

sub subClearPageFileatShutdown

end sub

sub subCheckAutoRun

end sub

sub subChangeAutoRun

end sub

sub subCheckPosixOS2

end sub

sub subRemovePosixOS2

end sub

sub subServerRole

end sub

sub subSpecificWebSettings

end sub

sub subSpecificSQLSettings

end sub

sub subSpecificDCSettings

end sub

sub subCheckServices

end sub

sub subDisableServices

end sub

sub subCheckHotfixes

end sub

sub subCheckNetBIOS

end sub

sub subDisableNetBIOS

end sub

sub subAllNICSTcpIpOnly

end sub

sub subAnyTcpIpFiltering

end sub

sub subCheckACLsOnDiags

end sub

sub subChangeACLsOnDiags

end sub

sub subCheckRegistrySecurity

end sub

sub subChangeRegistrySecurity

end sub

sub subBackupsDirACLs

end sub

sub subCheckAuditPolicy

	subMSG "TEST, Audit Policy"

	Dim objshell, oFSO , ofile
	Dim strLine, ArrNumber
	Dim strAudSysEvt, strAudLogEvt, strAudPriv, strAudPolChg, strAudAccMgmt, strAudProcess, strAudDSAccess, strAudAcctLogon, strCrashOnAuditF	
	Dim boolAudit

	boolAudit = TRUE

	set objshell = createobject ("wscript.shell")
	set oFSO = createobject ("scripting.filesystemobject")	

	objShell.run "secedit /export /cfg test.log /areas securitypolicy /verbose", 0, true
	
	if oFSO.fileexists ("test.log") = true then

		set oFile = oFSO.OpenTextFile ("test.log")

		Do While oFile.AtendOfStream = False
		
			strLine = oFile.Readline
			
			arrNumber = Split (strLine, "=")
			strLine = Left(strLine, 13)				

			Select Case strLine
			
				case "AuditSystemEv"
					StrAudSysEvt = arrNumber(1)
				case "AuditLogonEve"
					strAudLogEvt = arrnumber(1)
				case "AuditPrivileg"
					strAudPriv = arrnumber(1)
				case "AuditPolicyCh"
					strAudPolChg = arrnumber(1)
				case "AuditAccountM"
					strAudAccMgmt = arrnumber(1)
				case "AuditProcessT"
					strAudProcess = arrnumber(1)
				case "AuditDSAccess"
					strAudDSAccess = arrNumber(1)
				case "AuditAccountL"
					strAudAcctLogon = arrNumber(1)
				case "CrashOnAuditF"
					strCrashOnAuditf = arrNumber(1)
				case else
			
			End Select
		Loop

			ofile.close
			ofso.deletefile "test.log",true

			Select Case strAudSysEvt
				
				case 0
					'no auditing
					subMSG "WARNING, No Auditing on System Events"
					boolAudit = FALSE
				case 1,2
					'some auditing
					subMSG "NOTE, Some Auditing on System Events"
				case 3
					'all auditing
					subMSG "SECURE, Auditing on all System Events"
				case else
					'unknown
			end select

			Select Case strAudLogEvt
				
				case 0
					'no auditing
					subMSG "WARNING, No Auditing on Logon Events"
					boolAudit = FALSE
				case 1,2
					'some auditing
					subMSG "NOTE, Some Auditing on Logon Events"
				case 3
					'all auditing
					subMSG "SECURE, Auditing on all Logon Events"
				case else
					'unknown
			end select

			Select Case strAudPriv
				
				case 0
					'no auditing
					subMSG "WARNING, No Auditing on Privilege Use"
					boolAudit = FALSE
				case 1,2
					'some auditing
					subMSG "NOTE, Some Auditing on Privilege Use"
				case 3
					'all auditing
					subMSG "SECURE, Auditing on Privilege Use"
				case else
					'unknown
			end select

			Select Case strAudPolChg
				
				case 0
					'no auditing
					subMSG "WARNING, No Auditing on Policy Changes"
					boolAudit = FALSE
				case 1,2
					'some auditing
					subMSG "NOTE, Some Auditing on Policy Changes"
				case 3
					'all auditing
					subMSG "SECURE, Auditing on all Policy Changes"
				case else
					'unknown
			end select

			Select Case strAudAccMgmt
				
				case 0
					'no auditing
					subMSG "WARNING, No Auditing on Account Management"
					boolAudit = FALSE
				case 1,2
					'some auditing
					subMSG "NOTE, Some Auditing on Account Management"
				case 3
					'all auditing
					subMSG "SECURE, Auditing on all Account Management"
				case else
					'unknown
			end select

			Select Case strAudDSAccess
				
				case 0
					'no auditing
					subMSG "WARNING, No Auditing on Directory Services Access"
					boolAudit = FALSE
				case 1,2
					'some auditing
					subMSG "NOTE, Some Auditing on Directory Services Access"
				case 3
					'all auditing
					subMSG "SECURE, Auditing on all Directory Services Access"
				case else
					'unknown
			end select

			Select Case strAudAcctLogon
				
				case 0
					'no auditing
					subMSG "WARNING, No Auditing on Account Logons"
					boolAudit = FALSE
				case 1,2
					'some auditing
					subMSG "NOTE, Some Auditing on Account Logons"
				case 3
					'all auditing
					subMSG "SECURE, Auditing on all Account Logons"
				case else
					'unknown
			end select

			'Select Case strCrashonAuditF
			'	
			'	case 0
			'		'no auditing
			'	case 1,2
			'		'some auditing
			'	case 3
			'		'all auditing
			'	case else
			'		'unknown
			'end select

		if boolAudit = FALSE then
			subChangeAuditPolicy
		else
			subMSG "SECURE, Audit Policy OK"
		end if


	else
		subErrmsg "File test.log does not exist"
	end if
		
	set oFSO = nothing
	set objshell = nothing

end sub

sub subChangeAuditPolicy
	
	if boolAction = FALSE then
		'do nothing
	else
	
		subMsg "ACTION, Changing Password Policy"
	
		Dim oShell 
		Dim oFSO
		Dim strInf
		Dim oFile
		
		set oShell = createobject ("wscript.shell")
		set oFSO = createobject ("scripting.filesystemobject")
		
		if oFSO.fileexists ("new.inf") = true then
			oFSO.deletefile "new.inf", true
			subDebugMsg "new.inf found and deleted"
		else
		end if
		
		if oFSO.fileexists ("secedit.sdb") = true then
			subDebugMsg "deleting secedit.sdb"
			oFSO.deletefile "secedit.sdb", true
		else
		end if
				

		strInf = "[Version]" & vbcrlf & _
			"Signature=""$CHICAGO$""" & vbcrlf & _
			"Revision=1" & vbcrlf &_
			"[Profile Description]" & vbcrlf &_
			"Description=New Security Settings" & vbcrlf& _
			"[Event Audit]" & vbcrlf &_
			"AuditSystemEvents = 3" & vbcrlf &_
			"AuditLogonEvents = 3" & vbcrlf &_
			"AuditObjectAccess = 3" & vbcrlf &_
			"AuditPrivilegeUse = 3" & vbcrlf &_
			"AuditPolicyChange = 3" & vbcrlf &_
			"AuditAccountManage = 3" & vbcrlf &_
			"AuditProcessTracking = 3" & vbcrlf &_
			"AuditDSAccess = 3" & vbcrlf &_
			"AuditAccountLogon = 3" & vbcrlf &_
			"CrashOnAuditFull = 0" & vbcrlf

		set ofile = oFSO.OpenTextFile ("new.inf", 8, true)

		ofile.write strInf 
		subDebugMsg "new.inf written"
		ofile.close

		subDebugMsg "secedit started"
		oShell.run "secedit /configure /DB %systemroot%\security\secedit.sdb /cfg new.inf /verbose", 0, true
		
		subDebugMsg "secedit finished"
		
		if oFSO.fileexists  ("secedit.sdb") = true then
			subDebugMsg "deleting secedit.sdb"
			oFSO.deletefile "secedit.sdb", true
		else
		end if
		
		if oFSO.fileexists ("new.inf") = true then 
			subDebugMsg "deleting new.inf"
			oFSO.deletefile "new.inf", true
		else
		end if
		
		set oFSO = nothing
		set oShell = nothing
		
		subCheckAuditPolicy 

	end if

end sub

sub SubCheckModems

	Dim strComputer
	Dim objWMIService
	Dim colItems, objItem

	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_POTSModem",,48)
	
	For Each objItem in colItems
		
		If objItem.DeviceID <> "" then 
			subMsg "WARNING: Found " & objItem.Caption & ", to be secure disable it manually"
		else
			subMsg "SECURE: No modems found"
		end if
	Next

end sub

sub subCheckScreenSaverLockout
	
	subMsg "TEST, Do users have to enter password if screen saver is on"
	
	Dim oReg, oShell
	'Dim HKEY_USERS, HKEY_CURRENT_USER
	Dim strDefaultKeyPath, strCurrentKeyPath
	Dim arrNameValues, arrNameTypes, i
	Dim arrNameValues2, arrNameTypes2
	
	const HKEY_USERS = &H80000003
	const HKEY_CURRENT_USER = &H80000001

	set oReg = GetObject ("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	
	strDefaultKeyPath = ".DEFAULT\Control Panel\Desktop"
	strCurrentKeyPath = "Control Panel\Desktop"
	
	oReg.EnumValues HKEY_USERS, strDefaultKeyPath, arrNameValues, arrNameTypes
	subDebugMsg "Enumerating .DEFAULT"
	
	for i = 0 to Ubound (arrNameValues)
		
		if arrNameValues (i) = "ScreenSaverIsSecure" then
			subDebugMsg "Found ScreenSaverIsSecure value"
			
			set oShell = createobject ("wscript.shell")
			
			if oShell.regRead ("HKEY_USERS\.DEFAULT\Control Panel\Desktop\ScreenSaverIsSecure") = 0 then 
				subMsg "WARNING, Screen saver is not secure on new users"
				subChgScreenSaverLockout "HKEY_USERS\.DEFAULT\Control Panel\Desktop\ScreenSaverIsSecure"
			else 
				subMsg "SECURE, Screen saver is secure on new users"
			end if
			
			set oShell = nothing
			
		else 
			'subMsg "WARNING, Can not determine if lockout on Screensaver is enabled"
		end if
		
		
	next
	
	oReg.EnumValues HKEY_CURRENT_USER, strCurrentKeyPath, arrNameValues2, arrNameTypes2
	subDebugMsg "Enumerating Current User"
	for i = 0 to Ubound (arrNameValues2)
		if arrNameValues2 (i) = "ScreenSaverIsSecure" then
			
			subDebugMsg "Found ScreenSaverIsSecure value"
			
			set oShell = createobject ("wscript.shell")
			
			if oShell.regRead ("HKEY_CURRENT_USER\Control Panel\Desktop\ScreenSaverIsSecure") = 0 then 
				subMsg "WARNING, Screen saver is not secure on current user"
				subChgScreenSaverLockout "HKEY_CURRENT_USER\Control Panel\Desktop\ScreenSaverIsSecure"
			else 
				subMsg "SECURE, Screen saver is secure on current user"
			end if
			
			set oShell = nothing
			
		else 
			'subMsg "WARNING, Can not determine if lockout on Screensaver is enabled"
			
		end if
	next

	set oReg = nothing

end sub

sub subChgScreenSaverLockout (strReg)

	if boolAction = FALSE then
		'do nothing
	else
	
		subMsg "ACTION, Changing " & strReg & " to 1"

		Dim oShell

		set oShell = createobject ("wscript.shell")

		oShell.RegWrite strReg, "1"

		set oShell = nothing 

		subCheckScreenSaverLockout
	
	end if

end sub

sub subCheckNTFS
	
	Dim objFSO, objDrive
	Dim colDrives, boolFAT

	boolFAT = false

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set colDrives = objFSO.Drives
	For Each objDrive in colDrives

	 select case objDrive.DriveType

		case 0,1,3,4
			'do nothing
		case 2
			if objDrive.FileSystem = "FAT" then
				subMsg "WARNING, Drive " & objDrive.DriveLetter & ": is FAT"
				boolFAT = TRUE
			end if
		case 5
			if objDrive.FileSystem = "FAT" then
				subMsg "WARNING, Drive " & objDrive.DriveLetter & ": is FAT"
				boolFAT = TRUE
			end if			
			
		case else
			'do nothing
	    end select	

	Next
	
	if boolFAT = TRUE then
		subMsg "WARNING, At least one drive uses FAT, manually convert the partition and rerun script"
	else
		subMSG "SECURE, All drives are NTFS"
	end if
	
	set objFSO = nothing 
	
end sub

sub subCheckPwdPolicy 

	subMsg "TEST, Password Policy"

	Dim oShell, oFSO, oFile
	Dim strMinPwdAge, strMaxPwdAge, strMinPwdLen, strPwdCmplx, strline
	Dim strPwdHist, strLockoutBadCount, strReqLogon2Chg, strClearText
	Dim arrNumber, boolPwdSecure

	boolPwdSecure = TRUE

	set oShell = createObject ("wscript.shell")
	
	subDebugMsg "secedit started"
	
	oShell.run "secedit /export /cfg test.log /areas securitypolicy /verbose", 0, true
	
	subDebugMsg "secedit finished"
	
	set oFSO = createobject ("scripting.filesystemobject")
	
	if oFSO.fileexists ("test.log") = true then
		
		set oFile = oFSO.OpenTextFile ("test.log")
		
		Do while oFile.AtEndOfStream <> True
		
			strLine = oFile.Readline
			
			arrNumber = Split (strLine, "=")
			strLine = Left(strLine, 16)
			
			select case strLine				
										
				case "MinimumPasswordA"
					strMinPwdAge = arrNumber(1)
				case "MaximumPasswordA"
					strMaxPwdAge = arrNumber(1)
				case "MinimumPasswordL"
					strMinPwdLen = arrNumber(1)
				case "PasswordComplexi"
					strPwdCmplx = arrNumber(1)
				case "PasswordHistoryS"
					strPwdHist = arrNumber(1)
				case "LockoutBadCount "
					strLockoutBadCount = arrNumber(1)
				case "RequireLogonToCh"
					strReqLogon2Chg = arrNumber(1)
				case "ClearTextPasswor"
					strClearText = arrNumber(1)			
				case else
					'ignore all other lines
			end select
		Loop
		
		oFile.close
		
		if strMinPwdAge < 1 then 
			subMsg "WARNING, User can change password in timescales less than a day"
			boolPwdSecure = FALSE
		end if
		
		if strMaxPwdAge > 30 then
			subMsg "WARNING, Passwords expire in " & trim (strMaxPwdAge) & " days"
			boolPwdSecure = FALSE
		else
			subMsg "SECURE, Passwords expire in " & trim (strMaxPwdAge) & " days"
		end if
		
		if strMinPwdLen < 7 then 
			subMsg "WARNING, Passwords minimum length is to low"
			boolPwdSecure = FALSE
		else
			subMsg "SECURE, Passwords minimum length is acceptable"
		end if
		
		select case strPwdCmplx
		
			case 0
				subMsg "WARNING, Password Complexity requirements are disabled"	
				boolPwdSecure = FALSE
			case 1
				subMsg "SECURE, Password Complexity requirements are enabled"
			case else
				subMSG "WARNING, Can not determine Password complexity requirements"
				boolPwdSecure = FALSE
		end select
		
		if strPwdHist < 4 then
			subMsg "WARNING, Password history store is less than 5"
			boolPwdSecure = FALSE
		else
			subMsg "SECURE, Password history stores more than 5 passwords"
		end if
		
		if strLockoutBadCount > 0 then
			subMsg "SECURE, User account locks after " & trim (strLockoutBadCount) & " attempts" 
		else
			subMsg "WARNING, User Accounts never lockout"
			boolPwdSecure = FALSE
		end if 
		
		if boolPwdSecure = FALSE then
			subChgPwdPolicy
		end if
				
		'subDebugMsg "strMinPwdAge       : " & strMinPwdAge
		'subDebugMsg "strMaxPwdAge       : " & strMaxPwdAge
		'subDebugMsg "strMinPwdLen       : " & strMinPwdLen
		'subDebugMsg "strPwdCmplx        : " & strPwdCmplx
		'subDebugMsg "strPwdHist         : " & strPwdHist
		'subDebugMsg "strLockoutBadCount : " & strLockoutBadCount 
		'subDebugMsg "strReqLogon2Chg    : " & strReqLogon2chg
		'subDebugMsg "strClearText       : " & strClearText
		
		
	else
		subErrMsg "SecEdit failed to produce a report on security"
	end if
	
	if oFSO.fileexists ("test.log") = true then
		oFSO.deletefile "test.log", true
	end if
	
	set oShell = nothing
	set oFSO = nothing
end sub

sub subChgPwdPolicy

	if boolAction = FALSE then
		'do nothing
	else
	
		subMsg "ACTION, Changing Password Policy"
	
		Dim oShell 
		Dim oFSO
		Dim strInf
		Dim oFile
		
		set oShell = createobject ("wscript.shell")
		set oFSO = createobject ("scripting.filesystemobject")
		
		if oFSO.fileexists ("new.inf") = true then
			oFSO.deletefile "new.inf", true
			subDebugMsg "new.inf found and deleted"
		else
		end if
		
		if oFSO.fileexists ("secedit.sdb") = true then
			subDebugMsg "deleting secedit.sdb"
			oFSO.deletefile "secedit.sdb", true
		else
		end if
				

		strInf = "[Version]" & vbcrlf & _
			"Signature=""$CHICAGO$""" & vbcrlf & _
			"Revision=1" & vbcrlf &_
			"[Profile Description]" & vbcrlf &_
			"Description=New Security Settings" & vbcrlf& _
			"[System Access]" & vbcrlf &_
			"MinimumPasswordAge = 5" & vbcrlf &_
			"MaximumPasswordAge = 30" & vbcrlf& _
			"MinimumPasswordLength = 7" & vbcrlf& _
			"PasswordComplexity = 1" & vbcrlf &_
			"PasswordHistorySize = 5" & vbcrlf &_
			"LockoutBadCount = 3" & vbcrlf &_
			"RequireLogonToChangePassword = 0" & vbcrlf & _
			"ClearTextPassword = 0" & vbcrlf

		set ofile = oFSO.OpenTextFile ("new.inf", 8, true)

		ofile.write strInf 
		subDebugMsg "new.inf written"
		ofile.close

		subDebugMsg "secedit started"
		oShell.run "secedit /configure /DB %systemroot%\security\secedit.sdb /cfg new.inf /verbose", 0, true
		
		subDebugMsg "secedit finished"
		
		if oFSO.fileexists  ("secedit.sdb") = true then
			subDebugMsg "deleting secedit.sdb"
			oFSO.deletefile "secedit.sdb", true
		else
		end if
		
		if oFSO.fileexists ("new.inf") = true then 
			subDebugMsg "deleting new.inf"
			oFSO.deletefile "new.inf", true
		else
		end if
		
		set oFSO = nothing
		set oShell = nothing
		
		subCheckPwdPolicy 
	
	end if 

end sub

sub subCheckShares (strServer)

	Dim oWMI, colShares, oShare, boolAnyShares

	subMsg "TEST, Check Shares"
	
	Set oWMI = GetObject("winmgmts:" _
    		& "{impersonationLevel=impersonate}!\\" & strServer & "\root\cimv2")
	
	set colShares = oWMI.ExecQuery ("Select * from win32_share")
	
	boolAnyShares = FALSE
	For Each oShare in colShares
	
		if Instr (oShare.Name, "$") <> 0 then
			' hidden share found 
			subDebugMsg "Hidden share found"
			if oShare.Type = -2147483648 then
				subMsg "NOTE, Admin Hidden Share (" & oShare.Name & ")"
				'boolAnyShares = FALSE
			else
				boolAnyShares = TRUE
				subMsg "WARNING, Hidden Share has unusual permissions (" & oShare.Name & ")"
				subDisableShare strServer, oShare.Name
			end if
			
		else
			' no hidden shares
			boolAnyShares = TRUE
			subMsg "WARNING, Share Found (" & oShare.Name & ")"
			
			' Disable share
			subDisableShare strServer, oShare.Name
		end if 
	
	Next
	
	if boolShare = FALSE then
		subDisableAdminShares strLocal
	end if
	
	set oWMI = nothing
	
	if boolAnyShares = FALSE then 
		subMsg "SECURE, no non-Admin shares found"
		exit sub
	else
		subDebugmsg "boolAnyShares is TRUE"
	end if
	
	'subCheckShares strServer
	
end sub

sub subDisableShare (strServer, strShare)

	if boolAction = FALSE or strShare = "IPC$" then
		'do nothing
	else 
		subMsg "ACTION, Disabling Share " & strShare
		
		Dim oWMI, oShare, colShares
		
		Set oWMI = GetObject("winmgmts:" _
    			& "{impersonationLevel=impersonate}!\\" & strServer & "\root\cimv2")	
		
		Set colShares = oWMI.ExecQuery _
		    ("Select * from Win32_Share Where Name = '" & strShare & "'")
		For Each oShare in colShares
		
			oShare.Delete
		
		Next

		set oWMI = nothing
	
	end if

end sub

sub subDisableAdminShares (strServer)
	
	boolShare = TRUE
	
	if boolAction = FALSE then
		'do nothing
	else
		Dim oReg, oShell, strKeyPath
		Dim boolAdmin, arrValueNames, arrValueTypes, i
		
		subMsg "ACTION, Disabling Admin Shares"
		
		
		const HKEY_LOCAL_MACHINE = &H80000002
				
		
		Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
			strServer & "\root\default:StdRegProv")
		
		strKeyPath = "SYSTEM\CurrentControlSet\Services\lanmanserver\parameters"
		
		oReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes
		boolAdmin = FALSE
		For i=0 To UBound(arrValueNames)
		    if arrValueNames(i) = "AutoShareWks" or arrValueNames(i) = "AutoShareServer" then
		    	boolAdmin = TRUE
		    	subMsg "SECURE, System should have no admin shares (may need reboot)"
		    	set oReg = nothing
		    	exit sub
		    end if
		Next
		
		if boolAdmin = FALSE then
			
			subMsg "NOTE, Not found Reg entry. Inserting...."
			
			set oShell = createobject ("wscript.shell")
			
			oShell.regwrite "HKLM\" & strKeyPath & "\AutoshareWks", "0"
			oShell.regwrite "HKLM\" & strKeyPath & "\AutoShareServer", "0"
			
			boolAnyShares = FALSE
			
			set oShell = nothing
		end if

		subCheckShares strServer
	end if

end sub



sub subCheckAdministrator (strServer)
	
	
	subMsg  "TEST, Does Administrator exist"

	Dim User, oComputer, boolAdmin
	boolAdmin = ""
	Set oComputer = GetObject ("WinNT://" & strServer)
	
	oComputer.Filter = Array ("User")

	For each User in oComputer
		if User.Name = "Administrator" then
			
			boolAdmin = TRUE
			subMsg "NOTE, Administrator Exists"
			subCheckAdminGroup strLocal
			
		end if
	Next	

	set oComputer = nothing
	
	if boolAdmin = TRUE then
	else
		subMsg "SECURE, Administrator does not exist on this machine"
	end if
	
end sub

sub subCheckAdminGroup (strServer)

	
	subMsg "TEST, Is Administrator in the Administrators group"
	
	Dim Group, oUser, boolGroup
	boolGroup = ""
	Set oUser = GetObject ("WinNT://" & strServer & "/Administrator,user")
	
	For each Group in oUser.Groups
		if Group.Name = "Administrators" then
			
			boolGroup = TRUE
			subMsg "NOTE, Administrator is in the Administrators group"
			
			'here we can rename administrator and create a new one
			subRenameAndCreateAdmin strLocal
		end if
	Next

	set oUser = nothing
	
	if boolGroup = TRUE then
	else 
		subMsg "SECURE, Administrator user is a dummy User"
	end if
end sub

sub subRenameAndCreateAdmin (strServer)

	if boolAction = False then
		' do nothing
	else
		Dim oComputer
		Dim oUser
		Dim oMoveUser
		Dim oGroup
		
		subMsg "ACTION, renaming Administrator to AUSR_" & strServer
		Set oComputer = GetObject ("WinNT://" & strServer)
		set oUser = GetObject ("WinNT://" & strServer & "/Administrator,user")
		set oMoveUser = oComputer.MoveHere (oUser.ADsPath, "AUSR_" & strServer)
		
		set oUser = nothing
		set oMoveUser = nothing
		
		subCheckAdministrator strServer
		
		' Create Dummy Administrator user
		subMsg "ACTION, creating false Administrator"
		set oUser = oComputer.Create ("User", "Administrator")
		oUser.SetInfo
		oUser.AccountDisabled = -1
		oUser.SetInfo
		
		set oUser = nothing
		
		set oUser = GetObject ("WinNT://" & strServer & "/Administrator,user")
		set oGroup = GetObject ("WinNT://" & strServer & "/Guests,group")
		
		oGroup.Add (oUser.ADsPath)
		oGroup.Setinfo
		
		set oUser = nothing
		set oGroup = nothing
		
		subCheckAdministrator strServer
		
	end if

end sub



sub subCheckNoUsers (strServer)

	Dim oComputer, Group
		
	subMsg "TEST, Number of Users. Check numbers."
	
	set oComputer = GetObject ("WinNT://" & strServer)
	
	oComputer.Filter = Array ("Group")
	
	For each Group in oComputer
		subPullUsersFromGroup strServer, Group.Name
		'wscript.echo Group.Name
	Next
	
end sub

sub subPullUsersFromGroup (strServer, strGroup)

	Dim oGroup
	Dim User, noOfUsers
	
	Set oGroup = GetObject ("WinNT://" & strServer & "/" & strGroup & ",group")
	
	noOfUsers = 0
	For each User in oGroup.Members
		noOfUsers = noOfUsers + 1
	Next
	
	select case strGroup
	
		case "Administrators"
			subMsg "   " & strGroup & " contains " & noOfUsers & " members"
		case "Guests"
			subMsg "   " & strGroup & " contains " & noOfUsers & " members"
		case "Users"
			subMsg "   " & strGroup & " contains " & noOfUsers & " members"
		case else
	
	end select
		
		
	'subMsg strGroup & " contains " & noOfUsers & " members"
	
end sub

sub subCheckGuest (strServer)
	
	subMsg "TEST, Is Guest Account disabled?"
	
	'connect to localmachine, guest account
	subDebugMsg "Connecting to " & strServer & ", Guest User Account"
	Set oUser = GetObject ("WinNT://" & strServer & "/Guest,user")

	' determining status
	Select Case oUser.AccountDisabled

		Case -1
			'Account is disabled
			SubMsg "SECURE, Guest Account is disabled"
		Case 0
			'Account is enabled
			SubMsg "WARNING, Guest Account is enabled"
			subDisableGuest strLocal
		Case Else
			'we don;t know whats going on
			subErrMsg "Can not determine Guest account state"
	End Select
	
	'clean up objects
	set oUser = nothing

end sub

sub subDisableGuest (strServer)
	
	' We only want to perform actions if user has specified so.
	if boolAction = FALSE then
		' no action if user specified, no action
	else
		subMsg "ACTION, Disabling Guest Account"
		' connect using ADSI to guest user on localmachine
		subDebugMsg "Connecting to " & strServer & ", Guest User Account"
		Set oUser = GetObject ("WinNT://" & strServer & "/Guest,user")
		
		' set the account disabled field to -1
		subMsg "Setting Guest Account to Disabled"
		oUser.AccountDisabled = -1
		
		' action the changes
		oUser.Setinfo
		
		' clean up objects
		set oUser = nothing
		
		' check change worked
		subCheckGuest strServer	
	end if


end sub

sub subMsg (strMsg)
	
	wscript.echo Now & " Msg: " & strMsg

end sub

sub subDebugMsg (strMsg)
	if BoolDebug = TRUE then	
		'Dim strMsg
		wscript.echo Now & " Debug: " & strMsg
	else
	end if	
		
	
end sub

sub subErrMsg (strMsg)
	
	'Dim strMsg
	wscript.echo Now & " Err: " & strMsg
	wscript.quit(1)
end sub

sub SubDisplayUsage

	submsg	vbcrlf & vbcrlf & " Usage: chkSecurity.vbs [/?] [/debug] [/action]" & _
		vbcrlf & vbcrlf  & "    /?        - Displays this help" & _
		vbcrlf & "    /debug    - shows all debug output " & _
		vbcrlf & "    /action - displays warnings, and perform fixes" & _
		vbcrlf & vbcrlf & " Note: By default, script will only display" & _
		vbcrlf & " potential problems, it will not fix anything " & _
		vbcrlf & " unless the /action switch is specified." & vbcrlf

end sub