'   Recreates permissions on NT from a modified trustee.nlm log file.
' Must manually set the path (mass search replace the volume for c:\ or whatever)
' User user accounts or group from the left-most position of the user name.  must exist 
' Tries to recreate file system rights


strFileIn = InputBox("Input File",,"C:\script\Netware\trst04.txt")
strFileOut = InputBox("Log file",,"C:\perm.log")


  Set objShell = CreateObject("Wscript.Shell")
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  strTemp = objShell.ExpandEnvironmentStrings("%TEMP%")
  strTempFile = strTemp & "\PermResult.tmp"
  
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
	
' ex: "TRUSTEE","IT:\IM\HOME\MontgoJ","LONG","MontgoJ.OU.OU.DOMAIN","SRWCEMFA"
	strPath = cutQuote(arrIn(1))
	strUserArr = cutQuote(arrIn(3))
	strRights = cutQuote(arrIn(4))
	
	arrUser = split(strUserArr,".")
	strUser = arrUser(0)
	
	strF = false
	strC = false
	strW = false
	strR = false
					' R W C F  is possible with cacls
					' set the most predominate/important one per Netware Right.
		if( instr(strRights, "S")) then  strF = true
		if( instr(strRights, "R")) then  strR = true
		if( instr(strRights, "W")) then  strC = true
		if( instr(strRights, "C")) then  strC = true
		'if( instr(strRights, "E")) then  strCaclsRights = strCaclsRights & "E"
		if( instr(strRights, "M")) then  
			strR = true
			strC = true
		end if
		if( instr(strRights, "F")) then  strR = true
		'if( instr(strRights, "A")) then  strCaclsRights = strCaclsRights & "A"
	
	strCaclsRights = ""
	'if strW = true then strCaclsRights = "W"	' aparently can't set more than one right with cacls?
	if strR = true then strCaclsRights = "R"	' so, figure out the most important right
	if strC = true then strCaclsRights = "C"
	if strF = true then strCaclsRights = "F"
	
		strCommand = "%comspec% /c c:\xcacls.exe " & chr(34) & strPath & chr(34) & " /g " & chr(34) & "PAIRGAIN\" & strUser & chr(34) & ":" & strCaclsRights & " /E >> " & strTempFile
	'msgbox strCommand
	objOutFile.WriteLine strCommand
	objShell.run strCommand
	wscript.sleep 100  ' need to wait 1/10th second 
Loop

objOutFile.close
set objoutfile = nothing
set objInFile = nothing
set objFSO = nothing

function cutQuote(strCut)
	cutQuote = Mid(strCut, 2, len(strCut)-2)
end function