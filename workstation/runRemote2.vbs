strInFile = "c:\susloc.txt"
Set objFSO = CreateObject("Scripting.FileSystemObject") 
set objFile = objFSO.OpenTextFile(strInFile)
Do Until objFile.AtEndOfStream			
	strComputer= objFile.ReadLine	
	strResult = RemoteStartProcess("\\server\share\process.exe", strComputer)
	wscript.echo strResult
Loop

function RemoteStartProcess(strPath, strComputer)
	strCmd = "c:\data\pstools\psexec.exe \\" & strComputer & " " & chr(34) & strPath & chr(34)
	wscript.echo strCmd
	Set objShell = CreateObject("Wscript.Shell")
	set objExecObject = objShell.Exec(strCmd)
	
	start = now()	
	
	do while not objExecObject.StdOut.AtEndOfStream
		strText = objExecObject.StdOut.ReadLine()
		if instr(strText, "The command") >0 then
			RemoteStartProcess = strText
			exit Do
		end if
		if instr(strText, "could not start") >0 then
			RemoteStartProcess = "Couldn't start"
			exit do
		end if
		if dateDiff("s",Start,Now)>15 then 
			RemoteStartProcess = "Started but hung"
			exit do
		end if
	loop
end function