strComputer = inputbox("what computer are you reading scheduled jobs from?")
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colScheduledJobs = objWMIService.ExecQuery("SELECT * FROM Win32_ScheduledJob")
wscript.echo strComputer
For Each objJob in colScheduledJobs
    Wscript.Echo "Caption: " & objJob.Caption
    Wscript.Echo "Command: " & objJob.Command
    Wscript.Echo "Days Of Month: " & objJob.DaysOfMonth
    Wscript.Echo "Days Of Week: " & objJob.DaysOfWeek
    Wscript.Echo "Description: " & objJob.Description
    Wscript.Echo "Elapsed Time: " & objJob.ElapsedTime
    Wscript.Echo "Install Date: " & objJob.InstallDate
    Wscript.Echo "Interact with Desktop: " & objJob.InteractWithDesktop
    Wscript.Echo "Job ID: " & objJob.JobID
    Wscript.Echo "Job Status: " & objJob.JobStatus
    Wscript.Echo "Name: " & objJob.Name
    Wscript.Echo "Notify: " & objJob.Notify
    Wscript.Echo "Owner: " & objJob.Owner
    Wscript.Echo "Priority: " & objJob.Priority
    Wscript.Echo "Run Repeatedly: " & objJob.RunRepeatedly
    Wscript.Echo "Start Time: " & objJob.StartTime
    Wscript.Echo "Status: " & objJob.Status
    Wscript.Echo "Time Submitted: " & objJob.TimeSubmitted
    Wscript.Echo "Until Time: " & objJob.UntilTime
Next
