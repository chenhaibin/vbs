strComputer = "."

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colMonitoredProcesses = objWMIService. _        
    ExecNotificationQuery("select * from __instancecreationevent " _ 
        & " within 1 where TargetInstance isa 'Win32_Process'")

Set colMonitoredProcesses2 = objWMIService. _
   ExecNotificationQuery("select * from __instancedeletionevent " _ 
          & "within 1 where TargetInstance isa 'Win32_Process'")

i = 0
Do 'While i = 0

	wscript.echo "."

    Set objLatestProcess2 = colMonitoredProcesses2.NextEvent
   Wscript.Echo Now & " Process Deleted: " & objLatestProcess2.TargetInstance.Name	

    Set objLatestProcess = colMonitoredProcesses.NextEvent
    Wscript.Echo Now & " Process Created: " & objLatestProcess.TargetInstance.ExecutablePath & " (PID: " & objLatestProcess.TargetInstance.ProcessID &")"



Loop
