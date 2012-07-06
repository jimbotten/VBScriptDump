'On Error Resume Next

sNewUser = "newUser"
sPassword = "newpass"
sGroupname = "Administrators"

Set oWshNet = CreateObject("WScript.Network")
sComputerName = oWshNet.ComputerName
Set oComputer = GetObject("WinNT://" & sComputerName)
Set oUser = oComputer.Create("user", sNewUser)
oUser.SetPassword sPassword
oUser.Setinfo
wscript.echo "Created service account"
Set oGroup = GetObject("WinNT://" & sComputerName & "/" & sGroupname)
oGroup.Add(oUser.ADsPath)
wscript.echo "Added service account to group"

Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000
Const ADS_UF_PASSWD_CANT_CHANGE = &H40

Set usr = GetObject("WinNT://" & sComputerName & "/" & sNewUser)

oldFlags = usr.Get("UserFlags")
newFlags = oldFlags Or ADS_UF_DONT_EXPIRE_PASSWD
newFlags = newFlags Or ADS_UF_PASSWD_CANT_CHANGE
usr.Put "UserFlags", newFlags
usr.SetInfo
wscript.echo "Set user account properties"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	

strSvcName  = "wscsvc"

set objWMI = GetObject("winmgmts:\\" & sComputerName & "\root\cimv2")
set objService = objWMI.Get("Win32_Service.Name='" & strSvcName & "'")
intRC = objService.Change(,,,,,,sNewUser,sPassword)
if intRC > 0 then
   WScript.Echo "Error setting service account: " & intRC
else
   WScript.Echo "Successfully set service account"
end if