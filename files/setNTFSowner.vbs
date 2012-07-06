'   Recreates owners on NT from a modified trustee.nlm log file.
' Must manually set the path (mass search replace the volume for c:\ or whatever)
' User user accounts or group from the left-most position of the user name.  must exist 
' Tries to recreate file system ownership
' needs a c:\chown program


strFileIn = InputBox("Input File",,"C:\script\Netware\owner.txt")
strFileOut = InputBox("Log file",,"C:\own.log")
strDOMAIN = "domainName/"

  Set objShell = CreateObject("Wscript.Shell")
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  strTemp = objShell.ExpandEnvironmentStrings("%TEMP%")
  strTempFile = strTemp & "\OwnerResult.tmp"
  
  set objInFile = objFSO.OpenTextFile(strFileIn)
  set objOutFile = objFSO.CreateTextFile(strFileOut,True)  ' True = overwrite
  
  
  Dim objFile, strResults

  Const OpenAsDefault = -2
  Const FailIfNotExist = 0
  Const ForReading = 1

Do Until objInFile.AtEndOfStream
	strCaclsRights = ""
	strIn= objInFile.ReadLine
	arrIn = split(strIn, ",")
	

	strPath = cutQuote(arrIn(1))
	strUserArr = cutQuote(arrIn(3))
	
	arrUser = split(strUserArr,".")
	strUser = arrUser(0)
		
	strCommand = "%comspec% /c c:\chown.exe " & strDOMAIN & "\" & strUser & " " & chr(34) & strPath & chr(34) & " >> " & strTempFile
	'msgbox strCommand
	objOutFile.WriteLine strCommand
	objShell.run strCommand
	wscript.sleep 250  ' need to wait 1/10th second 
Loop

objOutFile.close
set objoutfile = nothing
set objInFile = nothing
set objFSO = nothing

function cutQuote(strCut)
	cutQuote = Mid(strCut, 2, len(strCut)-2)
end function