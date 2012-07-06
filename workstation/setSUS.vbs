Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = inputBox("Give a hostname to apply registry settings to")
strSUSServer = inputBox("Give a hostname for the SUS Server.  I.E. 'http://server.location.com'")
Err.Clear

set objRegProv = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer & "\root\default:StdRegProv")

strKeyPath = "SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
objRegProv.CreateKey HKEY_LOCAL_MACHINE,strKeyPath

objRegProv.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,"NoAutoUpdate",0
objRegProv.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,"RescheduleWaitTime",&H3c
objRegProv.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,"NoAutoRebootWithLoggedOnUsers",1
objRegProv.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,"AUOptions",4
objRegProv.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,"ScheduledInstallDay",0
objRegProv.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,"ScheduledInstallTime",5
objRegProv.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,"UseWUServer",1

strKeyPath = "SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
objRegProv.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"WUServer",strSUSServer
objRegProv.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,"WUStatusServer",strSUSServer


if (Err.Number <> 0) then
	MsgBox "Error: " & Err.Description
end if 

set objWMIService = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer & "\root\cimv2")
set colServices = objWMIService.ExecQuery ("SELECT * FROM win32_Service WHERE Name = 'wuauserv'")
for each objService in colServices
	errReturnCode = objService.StopService()
	wscript.sleep 15000 ' 15 seconds
	if errReturnCode = 0 then
		objService.StartService()	
	else
		msgbox ("Automatic Updates service could not stop.")
	end if
next

MsgBox("Script Finished")



