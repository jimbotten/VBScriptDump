Set objSearcher = CreateObject("Microsoft.Update.Searcher")

Wscript.Echo "Can automatically upgrade service: " & _
    objSearcher.CanAutomaticallyUpgradeService
Wscript.Echo "Client application ID: " & objSearcher.ClientApplicationID
Wscript.Echo "Online: " & objSearcher.Online
Wscript.Echo "Server selection: " & objSearcher.ServerSelection
Wscript.Echo "Service ID: " & objSearcher.ServiceID