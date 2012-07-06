On Error resume next
strInFile = inputbox("Where is the input file of computer names?","Input File Name","C:\comps.txt")  
strOutFile = inputbox("Where do you want the log file?","Log File Name","c:\SUSloc.txt")
const HKEY_LOCAL_MACHINE = &H80000002

Set objFSO = CreateObject("Scripting.FileSystemObject") 
set objFile = objFSO.OpenTextFile(strInFile)
set objOut = objFSO.CreateTextFile(strOutFile)
Do Until objFile.AtEndOfStream			
	strComputer= objFile.ReadLine	
	err.clear
	Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
	if err.number <> 0 then 
		objOut.WriteLine strComputer & vbTab & "Remote Connection Failure"
	else 
		strKeyPath = "SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
		strValueName = "WUServer"
		Return = objReg.GetExpandedStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
		If (Return = 0) And (Err.Number = 0) Then   
			objOut.WriteLine strComputer & vbTab & strValue
		Else
			objOut.WriteLine  strComputer & vbTab & "Registry Connection Failure. Error = " & Err.Number
		End If
	end if
Loop

objOut.Close
objFile.Close

set objOut = nothing
set objFile = nothing
set objFSO = nothing

wscript.echo "Script complete.  Check your log file.  " & strOutFile