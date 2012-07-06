' this script/function pings an IP range given and tells you the response, ip, and conenctivity

Set objExcel = WScript.CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add
objExcel.ActiveSheet.Name = "Shares"
objExcel.ActiveSheet.Range("A1").Activate
objExcel.ActiveCell.Value = "Computer"		'col header 1
objExcel.ActiveCell.Offset(0,1).Value = "Ping"
objExcel.ActiveCell.Offset(0,2).Value = "WMI Remote"
objExcel.ActiveCell.Offset(0,3).Value = "C Drive - B used,B Free"
objExcel.ActiveCell.Offset(0,4).Value = "D Drive - B used,B Free"
objExcel.ActiveCell.Offset(0,5).Value = "Total RAM"
objExcel.ActiveCell.Offset(0,6).Value = "Model"
objExcel.ActiveCell.Offset(0,7).Value = "OS"
objExcel.ActiveCell.Offset(0,8).Value = "MHz"
objExcel.ActiveCell.Offset(0,10).Value = "RAID Controller"
objExcel.ActiveCell.Offset(0,11).Value = "RAID ver"
objExcel.ActiveCell.Offset(0,12).Value = "BIOS ver"
objExcel.ActiveCell.Offset(0,13).Value = "Service Tag"

objExcel.ActiveCell.Offset(0,14).Value = "First NIC"
objExcel.ActiveCell.Offset(0,15).Value = "IP"
objExcel.ActiveCell.Offset(0,16).Value = "Subnet"
objExcel.ActiveCell.Offset(0,17).Value = "Gateway"
objExcel.ActiveCell.Offset(0,18).Value = "MAC"
objExcel.ActiveCell.Offset(0,19).Value = "DNS"
objExcel.ActiveCell.Offset(0,20).Value = "WINS"

objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down
' ******************************

On Error Resume Next

intRange = inputBox("Base B Class Range?",,"146.71")
int3rdStart = inputBox("What 3rd octet IP to start from?",,"214")
int3rdFinish = inputBox("What 3rd octet IP to end on",,"214")
int4thStart = inputBox("What 4th octet IP to start from",,"1")
int4thFinish = inputBox("What 4th octet IP to end on",,"255")

For q3 = int3rdStart to int3rdFinish
	For q4 = int4thStart to int4thFinish
		strComputer = intRange & "." & q3 & "." & q4
		
		objExcel.ActiveCell.Value = strComputer
		getINV(strComputer)
	Next 
Next 

set objExcel=nothing
msgbox("Script Finished")

 ' **********************
function GetINV(strComputer)
if pingComp(strComputer) then 
	objExcel.ActiveCell.Offset(0,1).Value = " Ping Success"
	Set System = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	if (Err.Number <> 0) then
			objExcel.ActiveCell.Offset(0,2).Value = "Error: " & Err.Description
			objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down			
	else
	 	'Err.Clear
		objExcel.ActiveCell.Offset(0,2).Value = "WMI Connect Success " 		
		objExcel.ActiveCell.Offset(0,3).Value = GetDiskSpace(System,"C")
		objExcel.ActiveCell.Offset(0,4).Value = GetDiskSpace(System,"D")
		
		
		Set colSettings = System.ExecQuery("SELECT * FROM Win32_ComputerSystem")
		For Each objComputer in colSettings
			objExcel.ActiveCell.Value = strComputer
			objExcel.ActiveCell.Offset(0,5).Value = objComputer.TotalPhysicalMemory
			objExcel.ActiveCell.Offset(0,6).Value = objComputer.Model
			objExcel.ActiveCell.Offset(0,7).Value = detectOS(System)
		Next
		
		Set colSettings = System.ExecQuery("SELECT * FROM Win32_Processor")
		
		intProcCount = 0
		For Each objProc in colSettings
			strClockspeed = objProc.CurrentClockSpeed
			intProcCount = intProcCount + 1
		Next
		
		objExcel.ActiveCell.Offset(0,8).Value = "'" & intProcCount & " - " & strClockspeed  & "MHz"
		
		set colSettings = nothing
			
		Set colSettings = System.ExecQuery("SELECT * FROM Win32_SCSIController")
		For Each SCSI in colSettings
			objExcel.ActiveCell.Offset(0,10).Value = SCSI.Caption
			objExcel.ActiveCell.Offset(0,11).Value = SCSI.Name
		Next		
		
		Set colSettings = System.ExecQuery("SELECT * FROM Win32_BIOS")
		For Each BIOS in colSettings
			objExcel.ActiveCell.Offset(0,12).Value = BIOS.SMBIOSBIOSVersion
			objExcel.ActiveCell.Offset(0,13).Value = BIOS.SerialNumber
		Next		
		
		GetNICs(System)
		
	objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down			
	end if 
else 
	
	objExcel.ActiveCell.Offset(0,1).Value = "Can't ping"
	objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down	
end if
end function

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
	
Function GetDiskSpace(objWMI, drive)
	SET colSettings = objWMI.ExecQuery("SELECT * FROM Win32_LogicalDisk")
	FOR EACH objDisk IN colSettings
		IF (objDisk.Name = drive & ":") THEN 
			GetDiskSpace = objDisk.FreeSpace
			GetDiskSpace = GetDiskSpace & " - " & objDisk.Size
		END IF
	NEXT
END FUNCTION

FUNCTION GetNICs( objWMI )
	SET colItems = objWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=TRUE",,48)
		
	FOR EACH objItem IN colItems
		strDesc = strDesc & objItem.Description & ","
		FOR EACH ip IN objItem.IPAddress 
			strIPAdd = strIPAdd & ip & ","
		NEXT
		FOR EACH ipSub IN objItem.IPSubnet
			strIPSub = strIPSub & ipSub & ","
		NEXT
		if isarray(objItem.DefaultIPGateway) then
			for each gateway in objItem.DefaultIPGateway
				strIPGate = strIPGate & gateway & ","
			next
			'msgbox ("Array!" & strIPGate)
		else
			strIPGate = objItem.DefaultIPGateway
			'msgbox ("no array" & strIPGate)
		end if
				
		strMAC = strMAC & objItem.MACAddress & ","
		
		if isArray(objItem.DNSServerSearchOrder) THEN
			for each dnsserv in objItem.DNSServerSearchOrder 
				strDNS = strDNS & dnsserv & ","
			next
		ELSE
			strDNS = objItem.DNSServerSearchOrder
		END IF
		
		'FOR EACH DNSstring IN objItem.DNSServerSearchOrder 
		'	strDNS = strDNS & DNSstring & ","
		'NEXT
		
		strWINS = strWINS & objItem.WINSPrimaryServer & "-" & objItem.WINSSecondaryServer & ","
	NEXT
	
	objExcel.ActiveCell.Offset(0,14).Value = strDesc
	objExcel.ActiveCell.Offset(0,15).Value = strIPAdd
	objExcel.ActiveCell.Offset(0,16).Value = strIPSub
	objExcel.ActiveCell.Offset(0,17).Value = strIPGate
	objExcel.ActiveCell.Offset(0,18).Value = strMAC
	objExcel.ActiveCell.Offset(0,19).Value = strDNS
	objExcel.ActiveCell.Offset(0,20).Value = strWINS
	
END FUNCTION

FUNCTION ListInstalledApps( strComputer )
	Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
	strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
	strEntry1a = "DisplayName"
	strEntry1b = "QuietDisplayName"
	strEntry2 = "InstallDate"
	strEntry3 = "VersionMajor"
	strEntry4 = "VersionMinor"
	strEntry5 = "EstimatedSize"

	Set objReg = GetObject("winmgmts://" & strComputer &  "/root/default:StdRegProv")
	objReg.EnumKey HKLM, strKey, arrSubkeys

	For Each strSubkey In arrSubkeys
		intRet1 = objReg.GetStringValue(HKLM, strKey & strSbkey, strEntry1a, strValue1)
		If intRet1 <> 0 Then
			objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1b, strValue1
		End If
		If strValue1 <> "" Then
			ListInstalledApps = ListInstalledApps & strValue1 & vbCrLf
		End If
	Next
	
END FUNCTION