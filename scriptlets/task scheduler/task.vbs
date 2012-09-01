    Set oScheduler = Server.CreateObject("Scheduler.SchedulingAgent.1")  
    Set objTask = oScheduler.Tasks.Add("mytest.job")

    objTask.ApplicationName = "D:\WINNT\system32\notepad.exe"
    objTask.Comment = "this is my test task"
    objTask.Creator = "db"
    objTask.WorkingDirectory = "D:\WINNT\system32"
    objTask.SetAccountInformation "domain\user","~Thursday"

    Set oScheduler = Nothing
    Set objTask = Nothing
