

''Wscript.Echo 	"# Name   : Process Log" & vbcrlf &_
		'"# Version: 1.0" & vbcrlf & _
		'"# Date   : " & Now & vbcrlf &vbcrlf '& _

On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_PerfProc_Process",,48)
For Each objItem in colItems
   
   'wscript.echo Now & "# " & objItem.Name & ";" & objItem.IDProcess & ";" & 
   Wscript.Echo "Caption: " & objItem.Caption
    'Wscript.Echo "CreatingProcessID: " & objItem.CreatingProcessID
    'Wscript.Echo "Description: " & objItem.Description
    'Wscript.Echo "ElapsedTime: " & objItem.ElapsedTime
    'Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    'Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    'Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    'Wscript.Echo "HandleCount: " & objItem.HandleCount
    Wscript.Echo "IDProcess: " & objItem.IDProcess
    'Wscript.Echo "IODataBytesPersec: " & objItem.IODataBytesPersec
    'Wscript.Echo "IODataOperationsPersec: " & objItem.IODataOperationsPersec
    'Wscript.Echo "IOOtherBytesPersec: " & objItem.IOOtherBytesPersec
    'Wscript.Echo "IOOtherOperationsPersec: " & objItem.IOOtherOperationsPersec
    'Wscript.Echo "IOReadBytesPersec: " & objItem.IOReadBytesPersec
    'Wscript.Echo "IOReadOperationsPersec: " & objItem.IOReadOperationsPersec
    'Wscript.Echo "IOWriteBytesPersec: " & objItem.IOWriteBytesPersec
    'Wscript.Echo "IOWriteOperationsPersec: " & objItem.IOWriteOperationsPersec
    Wscript.Echo "Name: " & objItem.Name
    'Wscript.Echo "PageFaultsPersec: " & objItem.PageFaultsPersec
    'Wscript.Echo "PageFileBytes: " & objItem.PageFileBytes
    'Wscript.Echo "PageFileBytesPeak: " & objItem.PageFileBytesPeak
    'Wscript.Echo "PercentPrivilegedTime: " & objItem.PercentPrivilegedTime
    'Wscript.Echo "PercentProcessorTime: " & objItem.PercentProcessorTime
    'Wscript.Echo "PercentUserTime: " & objItem.PercentUserTime
    'Wscript.Echo "PoolNonpagedBytes: " & objItem.PoolNonpagedBytes
    'Wscript.Echo "PoolPagedBytes: " & objItem.PoolPagedBytes
    'Wscript.Echo "PriorityBase: " & objItem.PriorityBase
    'Wscript.Echo "PrivateBytes: " & objItem.PrivateBytes
    'Wscript.Echo "ThreadCount: " & objItem.ThreadCount
    'Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    'Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    'Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
    'Wscript.Echo "VirtualBytes: " & objItem.VirtualBytes
    'Wscript.Echo "VirtualBytesPeak: " & objItem.VirtualBytesPeak
    'Wscript.Echo "WorkingSet: " & objItem.WorkingSet
    Wscript.Echo "WorkingSetPeak: " & objItem.WorkingSetPeak
   ' wscript.quit
Next