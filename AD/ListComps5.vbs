Set objExcel = WScript.CreateObject("Excel.Application")

strDomain = InputBox("Waht LDAP Container?", "LDAP", "OU=comps,DC=domain,DC=gov")

objExcel.Visible = True
objExcel.Workbooks.Add

objExcel.ActiveSheet.Name = "Computers"
objExcel.ActiveSheet.Range("A1").Activate
objExcel.ActiveCell.Value = "ComputerName"		'col header 1
objExcel.ActiveCell.Offset(0,1).Value = "adspath"	'col header 2

objExcel.ActiveCell.Offset(1,0).Activate		'move 1 down

set objConn = CreateObject("ADODB.Connection")
	objConn.Open "Provider=ADsDSOObject;"

Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConn

strQuery = "<LDAP://" & strDomain & ">;(objectclass=computer);name,adspath;subtree"

objCommand.CommandText = strQuery
wscript.echo strQuery

Set objRS = objCommand.Execute
	while not objRS.EOF
		objExcel.ActiveCell.Value = objRS.Fields("name")
		objExcel.ActiveCell.Offset(0,1).Value  = objRS.Fields("adspath")
		if not testEPACompName(objRS.Fields("name")) then 
			objExcel.ActiveCell.Offset(0,2).Value  = "Broken"
		end if
		objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down	
		objRS.MoveNext
	wend
    
objConn.close
set objConn = nothing
set objCommand = nothing
set objRS = nothing

msgbox "Done"

function testEPACompName(compname)
	testEPACompName = true
	
	'Desktop ANNNNAaaaaaaxxN
	'Server, workstation ANNNNAAAAANNN
	' four chars, starting at 2nd spot are numbers
	if not isnumeric(mid(compname,2,4)) then
		testEPACompName = false
	end if 
	' rightmost char is a number
	if not isnumeric(right(compname,1)) then
		testEPACompName = false
	end if 
	' from end of rpio to 3 from end must be chars
	if (len(compname)-3) > 0 then		
		for x = 6 to (len(compname)-3)
			if isnumeric(mid(compname,x,1)) then
				testEPACompName = false
				'msgbox compname & vbcrlf & "position " & x & " Broke it"
			end if
		next
	end if
'	msgbox(compname _
'		& vbCrLf & mid(compname,2,4) & vbCrLf & isnumeric(mid(compname,2,4)) _
'		& vbCrLf & right(compname,1) & vbCrLf & isnumeric(right(compname,1)))
	
end function