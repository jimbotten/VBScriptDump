on error resume next 
'strInputFile = inputbox("What is the path to the input file?","Input","C:\compo.txt")
'strComputer = inputbox("What computer?")

set objFSO = CreateObject("Scripting.FileSystemObject")
'set objFile = objFSO.OpenTextFile(strInputFile)

set objFile = objFSO.OpenTextFile("c:\compo.txt")
Do Until objFile.AtEndOfStream
	strComputer= objFile.ReadLine
	outline = strComputer
	if pingComp(strComputer) then 
'		wscript.echo "pinged " & strComputer
	
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
		if (objWMIService is not null) then
'			wscript.echo "Got " & strComputer
			Set colNets = objWMIService.ExecQuery ("Select * from Win32_PerfFormattedData_PerfNet_Server")
'			wscript.echo "Ran perfmon query on " & strComputer
			For Each objNet in colNets
					outline = outline & "|" & objNet.FilesOpen & "|" & objNet.LogonTotal
			Next
		end if 
	else
		'wscript.echo "no ping " & strComputer
	end if
	wscript.echo outline
Loop


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