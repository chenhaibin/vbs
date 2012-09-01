' FILE: iisinfo.vbs
' DESC: This displays information about the current IIS installation
'       and metabase.
' AUTH: Thomas L. Fredell
' DATE: 11/16/98
'

Option Explicit

On Error Resume Next

Dim oArgs, sServer, nInstance, sIISPath
Dim oIISSrv, sTmp

' First check args to ensure that a computer is specified
Set oArgs = Wscript.Arguments
If oArgs.Count <> 2 Then
	Wscript.Echo "USAGE:  iisinfo  [computer] [instance]"
	Wscript.Echo ""
	Wscript.Echo "        Displays information about the IIS server"
	Wscript.Echo "        instance located on the specified computer"
	Wscript.Echo "        using IIS administration objects."
	Wscript.Echo ""
	Wscript.Quit 1
End If
sServer = oArgs(0)
nInstance = oArgs(1)

' Next connect to the IIS server and dump details
' Note: IIS uses ADSI conventions, so we need to create the path to
'       the IIS server in the ADSI directory.
sIISPath = "IIS://" & sServer & "/w3svc/" & CInt(nInstance)
Set oIISSrv = GetObject(sIISPath)
If Err.Number <> 0 Then
	ShowErr "Unable to connect to IIS server instance using '" & _
		sIISPath & "'."
	Wscript.Exit Err
End If
sTmp = "iisinfo.vbs: Dumping configuration details for server " & _
	"at '" & sIISPath & "'"
Wscript.Echo sTmp
Wscript.Echo String(Len(sTmp), "-")
Wscript.Echo ""
DumpObject oIISSrv

' End of script
Wscript.Quit 0


Sub ShowErr(sDescription)
	Wscript.Echo "iisinfo.vbs: " & sDescription
	Wscript.Echo "iisinfo.vbs: An error occurred - code: " & _
		Hex(Err.Number) & " - " & Err.Description
End Sub


Sub DumpObject(o)
	DumpADSIProperties o
	Wscript.Echo "AccessExecute: " & o.AccessExecute
	Wscript.Echo "AccessFlags: " & o.AccessFlags
	Wscript.Echo "AccessNoRemoteExecute: " & o.AccessNoRemoteExecute
	Wscript.Echo "AccessNoRemoteRead: " & o.AccessNoRemoteRead
	Wscript.Echo "AccessNoRemoteScript: " & o.AccessNoRemoteScript
	Wscript.Echo "AccessNoRemoteWrite: " & o.AccessNoRemoteWrite
	Wscript.Echo "AccessRead: " & o.AccessRead
	Wscript.Echo "AccessScript: " & o.AccessScript
	Wscript.Echo "AccessSSL: " & o.AccessSSL
	Wscript.Echo "AccessSSL128: " & o.AccessSSL128
	Wscript.Echo "AccessSSLFlags: " & o.AccessSSLFlags
	Wscript.Echo "AccessSSLMapCert: " & o.AccessSSLMapCert
	Wscript.Echo "AccessSSLNegotiateCert: " & o.AccessSSLNegotiateCert
	Wscript.Echo "AccessSSLRequireCert: " & o.AccessSSLRequireCert
	Wscript.Echo "AccessWrite: " & o.AccessWrite
'
'	Wscript.Echo "AdminACL: " & CInt(o.AdminACL)
'	This won't work because AdminACL is a binary reference
'
	Wscript.Echo "AllowKeepAlive: " & o.AllowKeepAlive
	Wscript.Echo "AllowPathInfoForScriptMappings: " & o.AllowPathInfoForScriptMappings
	Wscript.Echo "AnonymousPasswordSync: " & o.AnonymousPasswordSync
	Wscript.Echo "AnonymousUserName: " & o.AnonymousUserName
	Wscript.Echo "AnonymousUserPass: " & o.AnonymousUserPass
	Wscript.Echo "AppAllowClientDebug: " & o.AppAllowClientDebug
	Wscript.Echo "AppAllowDebugging: " & o.AppAllowDebugging
	Wscript.Echo "AppFriendlyName: " & o.AppFriendlyName
	Wscript.Echo "AppIsolated: " & o.AppIsolated
	Wscript.Echo "AppOopRecoverLimit: " & o.AppOopRecoverLimit
	Wscript.Echo "AppPackageID: " & o.AppPackageID
	Wscript.Echo "AppPackageName: " & o.AppPackageName
	Wscript.Echo "AppRoot: " & o.AppRoot
	Wscript.Echo "AppWamClsid: " & o.AppWamClsid
	Wscript.Echo "AspAllowOutOfProcComponents: " & o.AspAllowOutOfProcComponents
	Wscript.Echo "AspAllowSessionState: " & o.AspAllowSessionState
	Wscript.Echo "AspBufferingOn: " & o.AspBufferingOn
	Wscript.Echo "AspCodepage: " & o.AspCodepage
	Wscript.Echo "AspEnableParentPaths: " & o.AspEnableParentPaths
	Wscript.Echo "AspExceptionCatchEnable: " & o.AspExceptionCatchEnable
	Wscript.Echo "AspLogErrorRequests: " & o.AspLogErrorRequests
	Wscript.Echo "AspMemFreeFactor: " & o.AspMemFreeFactor
	Wscript.Echo "AspQueueTimeout: " & o.AspQueueTimeout
	Wscript.Echo "AspScriptEngineCacheMax: " & o.AspScriptEngineCacheMax
	Wscript.Echo "AspScriptErrorMessage: " & o.AspScriptErrorMessage
	Wscript.Echo "AspScriptErrorSentToBrowser: " & o.AspScriptErrorSentToBrowser
	Wscript.Echo "AspScriptFileCacheSize: " & o.AspScriptFileCacheSize
	Wscript.Echo "AspScriptLanguage: " & o.AspScriptLanguage
	Wscript.Echo "AspScriptTimeout: " & o.AspScriptTimeout
	Wscript.Echo "AspSessionTimeout: " & o.AspSessionTimeout
	Wscript.Echo "AuthAnonymous: " & o.AuthAnonymous
	Wscript.Echo "AuthBasic: " & o.AuthBasic
	Wscript.Echo "AuthFlags: " & o.AuthFlags
	Wscript.Echo "AuthNTLM: " & o.AuthNTLM
	Wscript.Echo "AuthPersistence: " & o.AuthPersistence
	Wscript.Echo "CacheControlCustom: " & o.CacheControlCustom
	Wscript.Echo "CacheControlMaxAge: " & o.CacheControlMaxAge
	Wscript.Echo "CacheControlNoCache: " & o.CacheControlNoCache
	Wscript.Echo "CacheISAPI: " & o.CacheISAPI
	Wscript.Echo "CGITimeout: " & o.CGITimeout
	Wscript.Echo "ConnectionTimeout: " & o.ConnectionTimeout
	Wscript.Echo "CreateCGIWithNewConsole: " & o.CreateCGIWithNewConsole
	Wscript.Echo "CreateProcessAsUser: " & o.CreateProcessAsUser
	Wscript.Echo "DefaultDoc: " & o.DefaultDoc
	Wscript.Echo "DefaultDocFooter: " & o.DefaultDocFooter
	Wscript.Echo "DefaultLogonDomain: " & o.DefaultLogonDomain
	Wscript.Echo "DirBrowseFlags: " & o.DirBrowseFlags
	Wscript.Echo "DirBrowseShowDate: " & o.DirBrowseShowDate
	Wscript.Echo "DirBrowseShowExtension: " & o.DirBrowseShowExtension
	Wscript.Echo "DirBrowseShowLongDate: " & o.DirBrowseShowLongDate
	Wscript.Echo "DirBrowseShowSize: " & o.DirBrowseShowSize
	Wscript.Echo "DirBrowseShowTime: " & o.DirBrowseShowTime
	Wscript.Echo "DontLog: " & o.DontLog
	Wscript.Echo "EnableDefaultDoc: " & o.EnableDefaultDoc
	Wscript.Echo "EnableDirBrowsing: " & o.EnableDirBrowsing
	Wscript.Echo "EnableDocFooter: " & o.EnableDocFooter
	Wscript.Echo "EnableReverseDns: " & o.EnableReverseDns
	Wscript.Echo "FrontPageWeb: " & CStr(o.FrontPageWeb)
	Wscript.Echo "HttpCustomHeaders: " & Join(o.HttpCustomHeaders, ", ")
	Wscript.Echo "HttpErrors: " & Join(o.HttpErrors, ", ")
	Wscript.Echo "HttpExpires: " & o.HttpExpires
	Wscript.Echo "HttpPics: " & Join(o.HttpPics, ", ")
	Wscript.Echo "HttpRedirect: " & o.HttpRedirect
'
'	Wscript.Echo "IPSecurity: " & o.IPSecurity
'   We can't display this because it's a binary reference.
'
	Wscript.Echo "LogExtFileBytesRecv: " & o.LogExtFileBytesRecv
	Wscript.Echo "LogExtFileBytesSent: " & o.LogExtFileBytesSent
	Wscript.Echo "LogExtFileClientIp: " & o.LogExtFileClientIp
	Wscript.Echo "LogExtFileComputerName: " & o.LogExtFileComputerName
	Wscript.Echo "LogExtFileCookie: " & o.LogExtFileCookie
	Wscript.Echo "LogExtFileDate: " & o.LogExtFileDate
	Wscript.Echo "LogExtFileFlags: " & o.LogExtFileFlags
	Wscript.Echo "LogExtFileHttpStatus: " & o.LogExtFileHttpStatus
	Wscript.Echo "LogExtFileMethod: " & o.LogExtFileMethod
	Wscript.Echo "LogExtFileProtocolVersion: " & o.LogExtFileProtocolVersion
	Wscript.Echo "LogExtFileReferer: " & o.LogExtFileReferer
	Wscript.Echo "LogExtFileServerIp: " & o.LogExtFileServerIp
	Wscript.Echo "LogExtFileServerPort: " & o.LogExtFileServerPort
	Wscript.Echo "LogExtFileSiteName: " & o.LogExtFileSiteName
	Wscript.Echo "LogExtFileTime: " & o.LogExtFileTime
	Wscript.Echo "LogExtFileTimeTaken: " & o.LogExtFileTimeTaken
	Wscript.Echo "LogExtFileUriQuery: " & o.LogExtFileUriQuery
	Wscript.Echo "LogExtFileUriStem: " & o.LogExtFileUriStem
	Wscript.Echo "LogExtFileUserAgent: " & o.LogExtFileUserAgent
	Wscript.Echo "LogExtFileUserName: " & o.LogExtFileUserName
	Wscript.Echo "LogExtFileWin32Status: " & o.LogExtFileWin32Status
	Wscript.Echo "LogFileDirectory: " & o.LogFileDirectory
	Wscript.Echo "LogFilePeriod: " & o.LogFilePeriod
	Wscript.Echo "LogFileTruncateSize: " & o.LogFileTruncateSize
	Wscript.Echo "LogOdbcDataSource: " & o.LogOdbcDataSource
	Wscript.Echo "LogOdbcPassword: " & o.LogOdbcPassword
	Wscript.Echo "LogOdbcTableName: " & o.LogOdbcTableName
	Wscript.Echo "LogOdbcUserName: " & o.LogOdbcUserName
	Wscript.Echo "LogonMethod: " & o.LogonMethod
	Wscript.Echo "LogPluginClsId: " & o.LogPluginClsId
	Wscript.Echo "LogType: " & o.LogType
	Wscript.Echo "MaxBandwidth: " & o.MaxBandwidth
	Wscript.Echo "MaxBandwidthBlocked: " & o.MaxBandwidthBlocked
	Wscript.Echo "MaxConnections: " & o.MaxConnections
	Wscript.Echo "MaxEndpointConnections: " & o.MaxEndpointConnections
'
'	Wscript.Echo "MimeMap: " & o.MimeMap
'   This is a sub-object of IISMimeMap type, so we aren't going to dump the
'   contents
'
	Wscript.Echo "NetLogonWorkstation: " & o.NetLogonWorkstation
	Wscript.Echo "NTAuthenticationProviders: " & o.NTAuthenticationProviders
	Wscript.Echo "PasswordCacheTTL: " & o.PasswordCacheTTL
	Wscript.Echo "PasswordChangeFlags: " & o.PasswordChangeFlags
	Wscript.Echo "PasswordExpirePrenotifyDays: " & o.PasswordExpirePrenotifyDays
	Wscript.Echo "PoolIDCTimeout: " & o.PoolIDCTimeout
	Wscript.Echo "ProcessNTCRIfLoggedOn: " & o.ProcessNTCRIfLoggedOn
	Wscript.Echo "PutReadSize: " & o.PutReadSize
	Wscript.Echo "Realm: " & o.Realm
	Wscript.Echo "RedirectHeaders: " & o.RedirectHeaders
	Wscript.Echo "ScriptMaps: " & Join(o.ScriptMaps, ", ")
	Wscript.Echo "SecureBindings: " & Join(o.SecureBindings, ", ")
	Wscript.Echo "ServerAutoStart: " & o.ServerAutoStart
	Wscript.Echo "ServerBindings: " & Join(o.ServerBindings, ", ")
	Wscript.Echo "ServerComment: " & o.ServerComment
	Wscript.Echo "ServerListenBacklog: " & o.ServerListenBacklog
	Wscript.Echo "ServerListenTimeout: " & o.ServerListenTimeout
	Wscript.Echo "ServerSize: " & o.ServerSize
	Wscript.Echo "ServerState: " & o.ServerState
	Wscript.Echo "SSIExecDisable: " & o.SSIExecDisable
	Wscript.Echo "UNCAuthenticationPassthrough: " & o.UNCAuthenticationPassthrough
	Wscript.Echo "UploadReadAheadSize: " & o.UploadReadAheadSize
	Wscript.Echo "UseHostName: " & o.UseHostName
End Sub

Sub DumpADSIProperties(o)
	Wscript.Echo "Name: " & o.Name
	Wscript.Echo "ADsPath: " & o.ADsPath
	Wscript.Echo "Class: " & o.Class
	Wscript.Echo "GUID: " & o.GUID
	Wscript.Echo "Parent: " & o.Parent
	Wscript.Echo "Schema: " & o.Schema
End Sub