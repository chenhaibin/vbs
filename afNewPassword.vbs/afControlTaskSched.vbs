
' ****************************************
' Name: afControlTaskSched.vbs
' Author: Adrian Farnell
'
' Date: 10.12.2001
'
' Version: Alpha
'
' ****************************************

strjobname = "mytest.job"

' ** start Task Scheduler Object **
set ObjScheduler = wscript.createObject ("Scheduler.SchedulingAgent.1")
set objTask = objscheduler.Tasks.Add ( strjobname )

objTask.ApplicationName = "c:\windows\system32\notepad.exe"
objTask.Comment = "A Notepad for test"
objTask.Creator = "Adrian Farnell"
objTask.WorkingDirectory = "c:\windows\system32"
objTask.SetAccountInformation "ruth", "adr1an"


set objTask = objscheduler.tasks.item ( strjobname )
set objTrigger = objtask.triggers.add ()

objtrigger.begintime = now ()
objtrigger.flags = 0
objtrigger.duration = 0
strTriggertype = "daily"

select case strTriggerType
	case "daily"
		objtrigger.triggertype = 1
		objtrigger.daysinterval = 1 'every day
	case "weekly"
		objtrigger.triggertype = 2
		objtrigger.weeklyinterval = 1 'every week
		objtrigger.daysoftheweek = 2 'Monday
	case "monthly"
		objtrigger.triggertype = 3
		objtrigger.months = 4095 'every month
	case else
		objtrigger.triggertype = 0
end select

set objtrigger = nothing

set ObjScheduler = nothing
set objTask = nothing
