set oArgs = WScript.Arguments
set oNamed = oArgs.Named

strOut = oNamed("o")
strAll = oNamed("a")

strInputFile = oNamed("i")
strOutputFile = oNamed("o")
set objFSO = CreateObject("Scripting.FileSystemObject")
set objFile = objFSO.OpenTextFile(strInputFile)
set objFileOut = objFSO.CreateTextFile(strOutputFile)
Do Until objFile.AtEndOfStream
		arrLine = split(objFile.ReadLine, chr(9))
		objFileOut.Writeline join(arrLine, ",")		
Loop

objFileout.close
objfile.close
set objFSO = nothing