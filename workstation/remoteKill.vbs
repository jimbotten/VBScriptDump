REM Remote Pinger
REM 3/16/5 by Jim Montgomery
REM runs a ping to google.com on comptuers in comps.txt
rem pops up a msgbox while pinging...give it a sec!
rem then gives results out your command prompt, so run from cscript remoteping.vbs

set objFSO = CreateObject("Scripting.FileSystemObject")
set objFile = objFSO.OpenTextFile("comps.txt")
set objScript = CreateObject("Wscript.Shell")

Do Until objFile.AtEndOfStream
	strComputer= objFile.ReadLine
	strCommand = "cmd /c "& chr(34) &"c:\psshutdown.exe -r \\" & strComputer &  " >> c:\shutdown.log" & chr(34) 
	wscript.echo strCommand
	objScript.Run(strCommand)
	wscript.sleep 5000
Loop
set objFile = nothing
set objFSO = nothing

msgbox ("Script Complete")








