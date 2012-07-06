'	Script 3/16/05 by Jim Montgomery
'	This script takes an input of comps.txt which is a list of comptuer names, sperated by carriage returns
'	This script will {try to}connect to each machine name and query a list of its shares, and enumerate them in an excel spreadsheet.
'	Share types are: 0: Disk Drive, 1: Print Queue, -2147483645: IPC, -2147483648: Disk Drive
'	It will skip non pingable addresses, and give you errors on ones it can't manage




Set objExcel = WScript.CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add
objExcel.ActiveSheet.Name = "Shares"
objExcel.ActiveSheet.Range("A1").Activate
objExcel.ActiveCell.Value = "Computer"		'col header 1
objExcel.ActiveCell.Offset(0,1).Value = "ShareName"
objExcel.ActiveCell.Offset(0,2).Value = "SharePath"
objExcel.ActiveCell.Offset(0,3).Value = "ShareType"

objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down
' ******************************

On Error Resume Next

set objFSO = CreateObject("Scripting.FileSystemObject")
set objFile = objFSO.OpenTextFile("c:\comps.txt")
Do Until objFile.AtEndOfStream
	
strComputer= objFile.ReadLine
objExcel.ActiveCell.Value = strComputer

Err.Clear

if pingComp(strComputer) then 

	Set System = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")


	if (Err.Number <> 0) then
			objExcel.ActiveCell.Offset(0,1).Value = "Error: " & Err.Description
			objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down			
	else
	 	'Err.Clear
		Set colShares = System.ExecQuery("SELECT * FROM Win32_SHare")
		For Each objShare in colShares

			objExcel.ActiveCell.Value = strComputer
			objExcel.ActiveCell.Offset(0,1).Value = objShare.Name
			objExcel.ActiveCell.Offset(0,2).Value = objShare.Path
			objExcel.ActiveCell.Offset(0,3).Value = objShare.Type
			objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down	
		
		Next
		set colShares= nothing
	end if 
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