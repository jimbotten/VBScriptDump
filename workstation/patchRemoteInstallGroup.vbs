'*  remote Patch Installation, uses comps.txt to install specified patch remotely with specified parameters
'* 11/11/04  Jim Montgomery   updated 7/15/5


' ip = inputbox("Remote Hostname or IP?")
strPatchFileName = inputbox("Patch file name on your C drive?","Patch Name")  
'  Makes an input box where the user can enter a variable into strPatchFileName.  Window name is Patch Name
strParam = inputbox("Parameters to use with this patch","Paramaters","/uninstall /norestart") 
strTextOut = inputbox("File location to use as a log?","Log File","C:\patchlog.txt")
' Input box to enter parameters. Sets the window name to Parameters and the default entry is /uninstall /norestart
Set objFSO = CreateObject("Scripting.FileSystemObject")  ' Create an object we can use to open the text file holding all the computer names
set objFile = objFSO.OpenTextFile("comps.txt")		'  Open the text file with the handle objFile
set objOut = objFSO.CreateTextFile(strTextOut)
Do Until objFile.AtEndOfStream			' loop until we're at the end of that file
	strComputer= objFile.ReadLine		' read the entire current line into the strComputer variable (increments position in the file)
	if (IsConnectible(strComputer)) then	' run that computer name through the isConnectible function to make sure its on the network
		if(DoPatch(strComputer, strParam)) then		' if successful, run the DoPatch function to run the patch commands
			'msgbox "Success on " & strComputer	' if successful, spit out a message for each computer that it was successful.
			objOut.WriteLine "Success with patch " & strPatchFileName & " on " & strComputer
		else
			'msgbox "Patch copied but would not start on " & strComputer	' if patch fails to start, spit out message
			objOut.WriteLine "Patch " & strPatchFileName & " copied but would not start on " & strComputer
		end if
	else
		'msgbox strComputer & " not pingable"		' if the isConnectible fails, spit out error message
		objOut.WriteLine strComputer & " not pingable"
	end if
Loop		' go to next iteration

objOut.Close
objFile.Close
set objOut = nothing
set objFile = nothing	' this script is done, so free up the memory these objects use
set objFSO = nothing

msgbox ("Script Complete")  ' give a message so you know the whole process is complete


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
'      IsConnectible = "No connection - " & ip ' no TTL, its a dead link
      IsConenctible = 0
    Case Else
'      IsConnectible = "Connected - " & ip ' TTL means its live
      IsConnectible = 1
  End Select
End Function

Function DoPatch(ip, param)
	' returns 1 if the install starts, otherwise a 0
    Const OverwriteExisting = True		' Set a constant to make something readable later on
	
	' Set objFSO = CreateObject("Scripting.FileSystemObject")       ' I don't think we need this since that object should already exist when we run this function
	objFSO.CopyFile "C:\"& strPatchFileName, "\\"& ip & "\c$\" & strPatchFileName, OverwriteExisting	
	' this copies the patch from your c:\ folder to the remote c drive.   Note that you need admin permission to get to the hidden c$ share.
	
	set osvcRemote = GetObject("winmgmts:\\" & ip & "\root\cimv2")	' this creates an object representing the remote machine
	set oprocess = osvcRemote.Get("win32_process")	' this creates an object for the remote machine processes
	
	' ret = oprocess.create("c:\\" & strPatchFileName & " -q -z")  '* this is version 1 code for quiet mode, no reboot when done.
	ret = oprocess.create("c:\\" & strPatchFileName & " " & param)  ' this starts a process on the remote machine using c:\+patch.exe+params
	' you have to specify the path, and you have to use an escape character before backslash, which happens to be backslash.  So you get to backslashes to represent c:\
	if (ret <> 0) then			' if it spit out anything but a zero, it was a success
		bolSuccess=1
	else
		bolSuccess=1			' okay, I don't think I implemented this yet. It should give more info if the process dies or something
	end if
	set oprocess = nothing		' clean up objects at the end of the function
	set osvcRemote = nothing
	set objFSO = nothing
	DoPatch= bolSuccess
end function