'	Script 11/9/05 by Jim Montgomery
'	This script takes an input of comps.txt which is a list of comptuer names, sperated by carriage returns
'	This script will {try to}connect to each machine name and query it's physical RAM, and available disk space on the primary drive partition,
'	and enumerate them in an excel spreadsheet.
'	It will skip non pingable addresses, and give you errors on ones it can't manage

strFile = inputbox("What input file?", "Input File", "C:\comps.txt")

Set objExcel = WScript.CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add
objExcel.ActiveSheet.Name = "Inventory"
objExcel.ActiveSheet.Range("A1").Activate
objExcel.ActiveCell.Value = "Building Code"		'col header 1
objExcel.ActiveCell.Offset(0,1).Value = "Computer Name"
objExcel.ActiveCell.Offset(0,2).Value = "User Name"
objExcel.ActiveCell.Offset(0,3).Value = "Department"
objExcel.ActiveCell.Offset(0,4).Value = "Make"
objExcel.ActiveCell.Offset(0,5).Value = "Model"
objExcel.ActiveCell.Offset(0,6).Value = "CPU (MHz)"
objExcel.ActiveCell.Offset(0,7).Value = "RAM (MB)"
objExcel.ActiveCell.Offset(0,8).Value = "HDD (GB)"
objExcel.ActiveCell.Offset(0,9).Value = "Operating System"
objExcel.ActiveCell.Offset(0,10).Value = "Network Available"
objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down
' ******************************
On Error Resume Next
set objFSO = CreateObject("Scripting.FileSystemObject")
set objFile = objFSO.OpenTextFile(strFile)
Do Until objFile.AtEndOfStream
	
strComputer= objFile.ReadLine
objExcel.ActiveCell.Offset(0,1).Value = strComputer

Err.Clear

if pingComp(strComputer) then 

	Set System = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	if (Err.Number <> 0) then
			objExcel.ActiveCell.Offset(0,9).Value = "Error: " & Err.Description
			objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down			
	else
	 	'Err.Clear
		Set colCSes = System.ExecQuery("SELECT * FROM Win32_ComputerSystem")
		For Each objCS in colCSes
			objExcel.ActiveCell.Offset(0,2).Value = objCS.UserName
			objExcel.ActiveCell.Offset(0,4).Value = objCS.Manufacturer
			objExcel.ActiveCell.Offset(0,5).Value = objCS.Model
			objExcel.ActiveCell.Offset(0,7).Value = objCS.TotalPhysicalMemory/1048576
		Next
		Set colSettings = System.ExecQuery("SELECT * FROM Win32_Processor")
		For Each objProc in colSettings
			objExcel.ActiveCell.Offset(0,6).Value = objProc.CurrentClockSpeed
		Next	
		Set colSettings = System.ExecQuery("SELECT * FROM Win32_LogicalDisk")
		For Each objDisk in colSettings
			if objDisk.DeviceID="C:" then objExcel.ActiveCell.Offset(0,8).Value = objDisk.Size/1073741824
		Next
		Set colOperatingSystems = System.ExecQuery("Select * from Win32_OperatingSystem")
		For Each objOS in colOperatingSystems
			objExcel.ActiveCell.Offset(0,9).Value = objOS.Caption & " - " & _
			  objOS.ServicePackMajorVersion & "." & objOS.ServicePackMinorVersion
		Next
	objExcel.ActiveCell.Offset(0,10).Value = "Yep"
	objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down			
	end if 
else 
	objExcel.ActiveCell.Offset(0,10).Value = "Can't ping"
	objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down	
end if

Loop

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