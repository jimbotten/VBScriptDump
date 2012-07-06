'	Script 3/16/05 by Jim Montgomery
'	This script takes an input of comps.txt which is a list of comptuer names, sperated by carriage returns
'	This script will {try to}connect to each machine name and query it's physical RAM, and available disk space on the primary drive partition,
'	and enumerate them in an excel spreadsheet.
'	It will skip non pingable addresses, and give you errors on ones it can't manage

strFile = inputbox("What input file?", "Input", "c:\comps.txt")
Set objExcel = WScript.CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add
objExcel.ActiveSheet.Name = "Inventory"
objExcel.ActiveSheet.Range("A1").Activate
objExcel.ActiveCell.Value = "Computer"		'col header 1
objExcel.ActiveCell.Offset(0,1).Value = "DriveSpace"
objExcel.ActiveCell.Offset(0,2).Value = "Mem"
objExcel.ActiveCell.Offset(0,3).Value = "SystemName"
objExcel.ActiveCell.Offset(0,4).Value = "Model"
objExcel.ActiveCell.Offset(0,5).Value = "OS"
objExcel.ActiveCell.Offset(0,6).Value = "MHz"
objExcel.ActiveCell.Offset(0,7).Value = "MAC"
objExcel.ActiveCell.Offset(0,8).Value = "DNS"

'name
'IP
'subnet
'gateway
'dns servers
'wins servers
'mac
' network adapter
'server model
' bios level
' backplane firmware
' bmc version
' memory and slots
' procs (# and spd)
' controller
' service tag
' os

objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down
' ******************************
On Error Resume Next
set objFSO = CreateObject("Scripting.FileSystemObject")
set objFile = objFSO.OpenTextFile(strFile)
Do Until objFile.AtEndOfStream
	
strComputer= objFile.ReadLine
objExcel.ActiveCell.Value = strComputer

Err.Clear

if pingComp(strComputer) then 

	Set System = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	if (Err.Number <> 0) then
			objExcel.ActiveCell.Offset(0,5).Value = "Error: " & Err.Description
			objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down			
	else
	 	'Err.Clear
		objExcel.ActiveCell.Offset(0,1).Value = GetDiskSpace(System)
		
		
		Set colSettings = System.ExecQuery("SELECT * FROM Win32_ComputerSystem")
		For Each objComputer in colSettings
			objExcel.ActiveCell.Value = strComputer
			objExcel.ActiveCell.Offset(0,2).Value = objComputer.TotalPhysicalMemory
			objExcel.ActiveCell.Offset(0,3).Value = objComputer.Name
			objExcel.ActiveCell.Offset(0,4).Value = objComputer.Model
			objExcel.ActiveCell.Offset(0,5).Value = detectOS(System)
		Next
		set colSettings = nothing
		
		Set colSettings = System.ExecQuery("SELECT * FROM Win32_Processor")
		For Each objProc in colSettings
			objExcel.ActiveCell.Offset(0,6).Value = objProc.CurrentClockSpeed
		Next
		set colSettings = nothing		
	
		objExcel.ActiveCell.Offset(0,7).Value = GetMAC(System)
		objExcel.ActiveCell.Offset(0,8).Value = GetDNS(System)
		
	objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down			
	end if 
else 
	objExcel.ActiveCell.Value = strComputer
	objExcel.ActiveCell.Offset(0,5).Value = "Can't ping"
	objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down	
end if

Loop

set objWMIService = nothing
Set objFSO = nothing

msgbox("Script Finished")

 ' **********************
Function IsConnectible(strHost)
  ' Returns IP if strHost can be pinged, else 0
  Set objShell = CreateObject("Wscript.Shell")
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  strTemp = objShell.ExpandEnvironmentStrings("%TEMP%")
  strTempFile = strTemp & "\RunResult.tmp"

  Dim objFile, strResults

  intPings = 2
  intTO = 750

  Const OpenAsDefault = -2
  Const FailIfNotExist = 0
  Const ForReading = 1

  objShell.Run "%comspec% /c ping -n " & intPings & " -w " & intTO & " " & chr(34) & strHost & chr(34) & " >" & strTempFile, 0, True

  Set objFile = objFSO.OpenTextFile(strTempFile, ForReading, FailIfNotExist, OpenAsDefault)
  strResults = objFile.ReadAll
  objFile.Close

  if InStr(StrResults, "Unknown host") then 
	ip = "Unknown host"
  end if
  if instr(StrResults, "Destination host") then 
	ip = Right (StrResults, len(strResults) - InStr(strResults, "from"))	'everything right of the first "from"
	strpos = InStr(ip, ":")
	ip = Left (ip, Cint(strpos)-5)		'everything lef tof the :	
  end if 
  if InStr(StrResults, "[") then 
	ip = Right (StrResults, len(strResults) - InStr(strResults, "["))	'everything right of the first [
	strpos = InStr(ip, "]")
	ip = Left (ip, Cint(strpos)-1)				'everything left of the ]
  end if 

  Select Case InStr(strResults, "TTL=")
    Case 0
      IsConnectible = 0
    Case Else
      IsConnectible = 1
  End Select
End Function

function detectOS(osvcRemote)

     set oOSInfo = osvcRemote.InstancesOf("Win32_OperatingSystem")
     'Only one instance is ever returned (the currently active OS), even though the following is a foreach.
     for each objOperatingSystem in oOSInfo
		' msgbox(objOperatingSystem.OSType & " : " & objOperatingSystem.Version)
		detectOS = objOperatingSystem.Version
     next
     set oOSInfo = nothing

end function

function pingComp(strComputer)
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colPings = objWMIService.ExecQuery ("Select * From Win32_PingStatus where Address = '" & strComputer & "'")

    For Each objStatus in colPings
        If IsNull(objStatus.StatusCode) _
            or objStatus.StatusCode<>0 Then 
            pingComp = 0
        Else
            pingComp = 1
        End If
    Next
End function

Function GetMAC(objWMI)
	Set colItems = objWMI.ExecQuery( _
		"SELECT * FROM Win32_NetworkAdapter WHERE AdapterTypeId = 0 AND NetConnectionStatus = 2",,48)
	for each objItem in colItems
		GetMAC = objItem.MACAddress
	next
end function
	
Function GetDNS(objWMI)
	Set colConfigs = objWMI.ExecQuery( "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True",,48)
	For Each objConfig in colConfigs
		msgbox objConfig.caption
        for each serv in objNetCard.DNSServerSearchOrder
			msgbox serv
			strDNS = strDNS & " " & serv
		next
		GetDNS = strDNS
	Next
End Function

Function GetDiskSpace(objWMI)
	Set colSettings = objWMI.ExecQuery("SELECT * FROM Win32_LogicalDisk")
	For Each objDisk in colSettings
		if (objDisk.Name = "C:") then GetDiskSpace = objDisk.FreeSpace
	Next
end function
		

Function GetWINS(objWMI)
	const FULL_DNS_REGISTRATION = TRUE
	const DOMAIN_DNS_REGISTRATION = TRUE
	Set colItems = objWMI.ExecQuery( _
		"SELECT * FROM Win32_NetworkAdapter WHERE AdapterTypeId = 0 AND NetConnectionStatus = 2",,48)
	For Each objItem in colItems
		Set objShare = objWMI.Get("Win32_NetworkAdapterConfiguration")
		intSetSuffixes = objShare.SetDNSSuffixSearchOrder(arrNewDNSSuffixSearchOrder)
		
		Set colNetCards = objWMI.ExecQuery _
			("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
		For Each objNetCard in colNetCards
            objNetCard.SetDNSDomain(strNewDNSSuffix)
            objNetCard.SetDNSServerSearchOrder(arrDNS)
			objNetCard.SetWINSServer strWINS1, strWINS2
			objNetCard.SetDynamicDNSRegistration FULL_DNS_REGISTRATION, DOMAIN_DNS_REGISTRATION
			objNetCard.SetTCPIPNetBIOS(1)
		Next
	  Next
End Function