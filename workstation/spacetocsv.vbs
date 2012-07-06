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
		arrLine = split(objFile.ReadLine, "  ")
		objFileOut.Writeline Replace(join(arrLine, ","), chr(34), "")  
Loop

objFileout.close
objfile.close
set objFSO = nothing