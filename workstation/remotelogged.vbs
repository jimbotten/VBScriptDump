strInFile=inputbox("What is the input file path?")
strOutFile=inputbox("What is the output file path?")
set objFSO = CreateObject("Scripting.FileSystemObject")
set objFile = objFSO.OpenTextFile(strInFile)
set objFileOut = objFSO.CreateTextFile(strOutFile)
Do Until objFile.AtEndOfStream
	
	strComputer= objFile.ReadLine
	'wscript.echo strComputer & vbTab & PSLOG(strComputer)
	objFileOut.WriteLine strComputer & vbTab & PSLOG(strComputer)
	
Loop
wscript.echo "Done"

' =============================================================
   Function PSLOG (strServer)
' =============================================================
   Set oShell = WScript.CreateObject("WScript.Shell")
   fPSLOG = "c:\psloggedon.exe"
   strFullUser = " "
   strUserID = " "
   strUserDomain = " "
   Set oExec = oShell.Exec("cmd /c " & fPSLOG & " -l \\" & strServer & " | find /i ""\"" | find /i /v """" | find /i /v ""connecting To registry"" ")

   Do While Not oExec.StdOut.AtEndOfStream
      strText = oExec.StdOut.ReadLine()
         If InStr(strText, "\") > 0 Then
            'WScript.Echo "strText = " & strText
            'Get full user's domain
               myarray = Split(strText," ", 3)
               strFullUser = myarray(2)
               PSLOG = strFullUser 
         End If
   Loop
         
'If didn't find logged on user...

   If Eval(Len(strFullUser) < 3) = "True" Then
		PSLOG = "None"
   End If
   
End Function