Set objAutoUpdate = CreateObject("Microsoft.Update.AutoUpdate")
Wscript.Echo "Service enabled: " & objAutoUpdate.ServiceEnabled

Set objSettings = objAutoUpdate.Settings

Wscript.Echo "Notification level: " & objSettings.NotificationLevel
Wscript.Echo "Read-only: " & objSettings.ReadOnly
Wscript.Echo "Required: " & objSettings.Required
Wscript.Echo "Scheduled Installation Day: " & _
    objSettings.ScheduledInstallationDay
Wscript.Echo "Scheduled Installation Time: " & _
    objSettings.ScheduledInstallationTime