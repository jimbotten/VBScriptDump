strInFile = inputbox("What is the file name? ","INPUT","c:\netlogon.log")
'intCount = inputBox("How many lines per chunk?","LineCount","20000")
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Dim arrNetLogon()
Redim arrNetLogon(0)

Dim arrSubnets()
Redim arrSubnets(0)

'----Get Subnets from AD
set objRootDSE = GetObject("LDAP://RootDSE")
set objSitesCont = GetObject("LDAP://cn=Subnets,cn=sites," & objRootDSE.Get("configurationNamingContext") )
objSitesCont.Filter = Array("subnet")
for each objSubnet in objSitesCont
   addArray arrSubnets, objSubnet.Get("name") 
next
arrTmpArray = CombSort(arrSubnets)



set objFile = objFSO.OpenTextFile(strInFile)

'-----READ in the input file
Do Until objFile.AtEndOfStream			
	'wscript.echo ". "
	strLineArr = split(objFile.ReadLine, " ")
	if ubound(strLineArr) = 5 then
		strLine = IPtoHex(strLineArr(5)) & " " & strLineArr(4)
		if not searchArray(arrNetLogon,strLine) then
			addArray arrNetLogon,strLine
			'arrTmpArray = Bubble(arrNetLogon)
		end if
	end if 
	
	if ((ubound(arrNetLogon) mod 1000) = 0) and (ubound(arrNetlogon) > 0) then 
		'wscript.echo "Completed 1000 lines in array." & vbtab & Time
	end if
Loop
wscript.echo "Finished creating Array"
arrTmpArray = CombSort(arrNetLogon)
'x=1	
'lineCount=0
set objOutFile = objFSO.CreateTextFile(strInFile & ".OUTPUT.txt")

for each strLine in arrNetLogon
	arrout = split(strLine," ")
	if ubound(arrout) =1 then   ' if this line has two items in it then
		'if searchArray(arrsubnet,HextoIP(arrout(0))) then
		'	msgbox("found " & HextoIP(arrout(0)) & " in the subnet
		'end if
		objOutFile.WriteLine HextoIP(arrout(0)) & vbtab & arrout(1)	
	end if
	'lineCount=lineCount+1
	'if lineCount>cint(intCount) then
	'	objOutFile.close
	'	x=x+1
	'	set objOutFile = objFSO.CreateTextFile(strInFile & x & "OUTPUT.txt")
	'	lineCount=0
	'	wscript.echo "Moving to " & strInFile & x & "OUTPUT.txt" & vbTab & "Finished " & intCount & " lines of original file"
	'end if
next

objFile.close
objOutFile.close

function Bubble(arrNames)
	For i = (UBound(arrNames) - 1) to 0 Step -1
		For j= 0 to i
			If UCase(arrNames(j)) > UCase(arrNames(j+1)) Then
				'wscript.echo arrnames(j) & " > " & arrnames(j+1)
				strHolder = arrNames(j+1)
				arrNames(j+1) = arrNames(j)
				arrNames(j) = strHolder
			End If
		Next
	Next
	Bubble = arrNames
End function

Function CombSort(arrNames)
	Dim i,j,gap,x,OK
	  
	gap = ubound(arrNames)
	OK = True
	While OK
		'You can try other values, but 1.33 seems to be the best
		gap= Int(gap/1.33)
		If gap < 1 Then gap = 1
		OK = (gap <> 1)
		For i = 0 To ubound(arrNames) - gap
			j = i + gap
			If arrNames(i) > arrNames(j) Then
				x = arrNames(i)
				arrNames(i) = arrNames(j)
				arrNames(j) = x
				OK = True
			End If
		Next
	Wend
	CombSort = arrNames
End Function

sub addArray(arrArray, strText)
	intTop = Ubound(arrArray)
	redim preserve arrArray(intTop+1)
	arrArray(intTop+1)=strText			
end sub

function searchArray(arrArray, strText)
	searchArray = False
	for each strEntry in arrArray
		if strText = strEntry then searchArray = True
	next
end function

function IPtoHex(strIP)
	on error resume next 
	Dim strLen
	Dim intPos, intPos2, intPos3, intPos4
	Dim strFirst, strSecond, strThird, strFourth

	strLen = len(strIP)
	intPos = instr(strIP,".")
	intPos2 = instr(intPos+1,strIP,".")
	intPos3 = instr(intPos2+1,strIP,".")

	strFirst = Mid(strIP, 1, intPos-1)
	strSecond = Mid(strIP, intPos+1, intPos2-intPos-1)
	strThird = Mid(strIP, intPos2+1, intPos3-intPos2-1)
	strFourth = Mid(strIP, intPos3+1, strLen-intPos3)

	IPtoHex = hex(strFirst) & "." & Hex(strSecond) & "." & Hex(strThird) & "." & Hex(strFourth)
	If Err.Number Then
        msgbox err.number & err.message
        Err.Clear
    End If
End Function

Function HexToDec(sHexStr)
            HexToDec = CLng("&H" & Trim(sHexStr))
End Function

function  HextoIP(strHex)
	arrhex = split (strHex, ".")
	if ubound(arrhex) = 3 then
		strFirst = HexToDec(arrhex(0))
		strSecond = HexToDec(arrhex(1))
		strThird = HexToDec(arrhex(2))
		strFourth = HexToDec(arrhex(3))
	
	end if
	HextoIP = strFirst & "." & strSecond  & "." & strThird & "." & strFourth
end function

