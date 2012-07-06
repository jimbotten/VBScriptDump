' this script/function pings an IP range given and tells you the response, ip, and conenctivity

Set objExcel = WScript.CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add
objExcel.ActiveSheet.Name = "Shares"
objExcel.ActiveSheet.Range("A1").Activate
objExcel.ActiveCell.Value = "Computer"		'col header 1
objExcel.ActiveCell.Offset(0,1).Value = "ShareName"

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
		objExcel.ActiveCell.Offset(0,1).Value = 	Isconnectible(intRange & "." & q3 & "." & q4)
		objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down
	Next 
Next 

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