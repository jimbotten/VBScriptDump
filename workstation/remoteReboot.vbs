'  Remote Shutdown  by Jim Montgomery
On error resume next

strInFile = inputbox("What is the input file")

set objFSO = CreateObject("Scripting.FileSystemObject")
set objFile = objFSO.OpenTextFile(strInFile)
Do Until objFile.AtEndOfStream
	strComputer= objFile.ReadLine

	set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	set colOS = objWMIService.ExecQuery ("SELECT * FROM Win32_OperatingSystem")
	err.clear
	
	for each objOS in colOS
		objOS.Reboot()
	next
	if err.number >0 then 
		wscript.echo "Error, but rebooting"
	else
		wscript.echo "Rebooting " & strComputer
	end if
loop
set colOS = nothing
set objWMIService = nothing
