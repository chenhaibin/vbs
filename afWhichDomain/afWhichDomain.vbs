On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48)
For Each objItem in colItems

    Wscript.Echo "Domain: " & objItem.Domain
    Wscript.Echo "DomainRole: " & objItem.DomainRole

    Wscript.Echo "UserName: " & objItem.UserName
    Wscript.Echo "Workgroup: " & objItem.Workgroup

Next