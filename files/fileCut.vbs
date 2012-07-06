set objFSO = CreateObject("Scripting.FileSystemObject")
strFileName = inputBox("What file to cut up?","FileName","c:\netlogon.log")
intCount = inputBox("How many lines per chunk?","LineCount","20000")

set objFile = objFSO.OpenTextFile(strFileName)
x = 1
y = 0
set objOUTFile = objFSO.CreateTextFile(strFileName & cstr(x) & ".CHOP.txt")

Do Until objFile.AtEndOfStream
	strLine = objFile.ReadLine
	objOUTFile.writeline(strLine)
	y=y+1
	'wscript.echo "y = " & y & vbTab & intCount
	if (y >= cint(intCount)) then 
		objOUTFile.Close
		x=x+1
		y=0
		'msgbox (strFileName & cstr(x) & ".CHOP.txt")
		set objOUTFile = objFSO.CreateTextFile(strFileName & cstr(x) & ".CHOP.txt")
	end if
Loop
objFile.Close
objOUTFile.Close
