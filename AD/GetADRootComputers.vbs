' GetADRootComputers.vbs
' by Jim Montgomery
' This script asks for a AD container and a DC name.  It will get a list of computer objects in that container, 
' and create a spreadsheet of their names, IP addresses, when they last set computer password with the domain, and if they are connectible.  HOT!

Set objShell = CreateObject("Wscript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objExcel = WScript.CreateObject("Excel.Application")

strTemp = objShell.ExpandEnvironmentStrings("%TEMP%")
strTempFile = strTemp & "\RunResult.tmp"
ipaddy="blank"

strDN = InputBox("Enter the distinguished name of a container", "Container", "CN=Computers,DC=domain,DC=com")
strDomainC = InputBox("Enter the hostname of an AD domain controller", "DC")


objExcel.Visible = True
objExcel.Workbooks.Add

objExcel.ActiveSheet.Name = "Computers"
objExcel.ActiveSheet.Range("A1").Activate
objExcel.ActiveCell.Value = "ComputerName"		'col header 1
objExcel.ActiveCell.Offset(0,1).Value = "pwdLastSet"	'col header 2
objExcel.ActiveCell.Offset(0,2).Value = "Connectible"	'col header 3
objExcel.ActiveCell.Offset(0,3).Value = "IP"	'col header 3

objExcel.ActiveCell.Offset(1,0).Activate		'move 1 down

Set objContainer = GetObject( "LDAP://" & strDomainC & "/" & strDN )

objContainer.Filter = Array("computer")
For Each objChild In objContainer
	ipaddy = "blank"
	CurrentComputerName = objChild.CN
	objExcel.ActiveCell.Value = CurrentComputerName

  
	set objDate = objChild.Get("pwdLastSet")

	objExcel.ActiveCell.Offset(0,1).Value = LastSet(objDate)
	objExcel.ActiveCell.Offset(0,2).Value = IsConnectible(CStr(CurrentCOmputerName), ipaddy)
	objExcel.ActiveCell.Offset(0,3).Value = ipaddy
	objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down
Next

set objShell= nothing
set objFSO = nothing

MsgBox ("Script Complete")

Function LastSet(objDate)
  lngHigh = objDate.HighPart
  lngLow = objDate.LowPart

  If (lngHigh = 0) And (lngLow = 0) Then
    lngAdjust = 0
  End If

  lngDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
    + lngLow) / 600000000 - lngAdjust) / 1440

  LastSet= CDate(lngDate)
End Function




'Function IsConnectible(strHost, intPings, intTO)
Function IsConnectible(strHost,ip)
' Returns True if strHost can be pinged.
' Based on a program by Alex Angelopoulos and Torgeir Bakken.
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

'#############################3            get the IP address!!!

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
      IsConnectible = False
    Case Else
      IsConnectible = True
  End Select
End Function