'	GetDescription Script 10/13/06 by Jim Montgomery
'	This script takes an input of comps.txt which is a list of comptuer names, sperated by carriage returns
'	This script will ping each computer, then {try to}connect to each machine name and query it's computer description with WMI
'	and then plop them in an excel spreadsheet.
'	It will skip non pingable addresses, and give you errors on ones it can't manage

strFile = inputbox("What input file will supply computer names?" & vbCrLf & "One computer per line.", "Input", "c:\comps.txt")
Set objExcel = WScript.CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add
objExcel.ActiveSheet.Name = "Inventory"
objExcel.ActiveSheet.Range("A1").Activate
objExcel.ActiveCell.Value = "Computer"		'col header 1
objExcel.ActiveCell.Offset(0,1).Value = "Desc"

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
			objExcel.ActiveCell.Offset(0,1).Value = "Error: " & Err.Description
			objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down			
	else
		Set colSettings = System.ExecQuery("SELECT * FROM Win32_OperatingSystem")
		For Each objComputer in colSettings
			objExcel.ActiveCell.Value = strComputer
			objExcel.ActiveCell.Offset(0,1).Value = objComputer.Description
			'objExcel.ActiveCell.Offset(0,2).Value = objComputer.Caption
		Next
		set colSettings = nothing
		
	objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down			
	end if 
else 
	objExcel.ActiveCell.Value = strComputer
	objExcel.ActiveCell.Offset(0,1).Value = "Can't ping"
	objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down	
end if

Loop

set objWMIService = nothing
Set objFSO = nothing

msgbox("Script Finished")

 ' **********************

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