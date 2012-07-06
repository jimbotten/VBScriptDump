' this script/function pings a name given from comps.txt and tells you the response, ip, and conenctivity

strServiceName = "lanmanserver"
strInputFile = inputbox("What is the path to the input file?","Input","C:\comps.txt")
Set objExcel = WScript.CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add
objExcel.ActiveSheet.Name = "Shares"
objExcel.ActiveSheet.Range("A1").Activate
objExcel.ActiveCell.Value = "Computer"		'col header 1
objExcel.ActiveCell.Offset(0,1).Value = "IP"
objExcel.ActiveCell.Offset(0,2).Value = "Server Service State"

objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down

' ******************************

On Error Resume Next

set objFSO = CreateObject("Scripting.FileSystemObject")
set objFile = objFSO.OpenTextFile(strInputFile)
Do Until objFile.AtEndOfStream
	strComputer= objFile.ReadLine
	objExcel.ActiveCell.Value = strComputer
	objExcel.ActiveCell.Offset(0,1).Value = 	Isconnectible(strComputer)
	objExcel.ActiveCell.Offset(0,2).Value = 	GetServiceState(strServiceName, strComputer)
	objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down
Loop

set objFile = nothing
Set objFSO = nothing
set objExcel=nothing
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
      IsConnectible = "No connection - " & ip ' no TTL, its a dead link
    Case Else
      IsConnectible = "Connected - " & ip ' TTL means its live
  End Select
End Function

Function GetServiceState(strServiceName, strComputer)
 	' requires DetectOS function
	ON error resume next
	Err.Clear
   	set osvcRemote = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	if (Err.Number <> 0) then
		GetServiceState= "Can't Connect: " & Err.Description
	else
    	GetServiceState = "Service Not Found"
		
		osType = DetectOS(osvcRemote) 
		'msgbox osType & " and left 3 is " & left(osType,3)
		strType = left(osType,3)
		if (instr(strType, "W2K") or instr(strType, "WXP") or instr(strType, "2K3"))then		
			Set colServices = osvcRemote.ExecQuery ("Select * from Win32_Service where Name='" & strServiceName & "'")
			'msgbox "count = " & colServices.count
			For Each objService in colServices
				GetServiceState = objService.State
			Next
		end if
	end if
	set osvcRemote = nothing
	
end function


function detectOS(osvcRemote)

     set oOSInfo = osvcRemote.InstancesOf("Win32_OperatingSystem")
     'Only one instance is ever returned (the currently active OS), even though the following is a foreach.
     for each objOperatingSystem in oOSInfo

          if (objOperatingSystem.OSType <> 18) then
               ' Make sure that this computer is Windows NT-based.
               systemType = "OLD"
          else
               if (objOperatingSystem.Version = "5.0.2195") then
                    ' Windows 2000 SP2, SP3, SP4.
                    if (objOperatingSystem.ServicePackMajorVersion = 2) then 
                    	systemType = "OLD"		' SP2 can't be sussed, therefore its old
                    end if
                    if (objOperatingSystem.ServicePackMajorVersion = 3) then 
                    	systemType = "W2KSP3"
                    end if
                    if (objOperatingSystem.ServicePackMajorVersion = 4) then
                        systemType = "W2KSP4"
                    end if

               elseif (objOperatingSystem.Version = "5.1.2600") then
                    ' Windows XP RTM, SP1.					
                    if (objOperatingSystem.ServicePackMajorVersion = 0) or (objOperatingSystem.ServicePackMajorVersion = 1) then
                         systemType = "WXPSP1"
                    end if
					if (objOperatingSystem.ServicePackMajorVersion = 2) then
						 systemType = "WXPSP2"
					end if

               elseif (objOperatingSystem.Version = "5.2.3790") then
                    ' Windows Server 2003 RTM
                    if (objOperatingSystem.ServicePackMajorVersion = 0) then
                         systemType = "2K3SP0"
                    end if
               end if

               if (systemType = "") then
                    'This was a Windows NT-based computer, but not with a valid service pack.
                    systemType = "OLD"
               end if
          end if

     next

     detectOS = systemType

end function