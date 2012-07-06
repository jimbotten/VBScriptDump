'  Edit these things:
Const EMAIL_FROM = "email@com.com"
Const EMAIL_TO = "email@com.com"
Const EMAIL_SUBJECT_LINE = "Automated AD Replication Report"

' ReadRepl5.vbs
' This script will search for all domain controllers from its current domain, and 
' run 'repadmin /showreps servername' on each of them, and send an email with results when complete
'  You can specify an output log file with /o:OutfileName.txt on the command line (/o for OUTPUT)
' By defualt it shows only failed computers.  You can show all output with /a:Y  (/a for ALL)
' You can schedule this to run with at

set oArgs = WScript.Arguments
set oNamed = oArgs.Named

strOut = oNamed("o")
strAll = oNamed("a")
' ("Where does the output file go?" & vbCrLf & "(Use /o: switch)" & vbCrLf & _
'		"e.g.: readrepl /o:c:\out.txt","Out", "C:\out.txt")

set objFSO = CreateObject("Scripting.FileSystemObject")
if strOut <> "" then set objOutput = objFSO.CreateTextFile(strOut)

arrDCs = getControllers

strLog = ""

for each strComputer in arrDCs
	'wscript.echo strComputer
	arrOut = ReadRepl(strComputer)
	if strOut <> "" then PrintReport strComputer, arrOut, objOutput
	if strAll = "Y" then 
		PrepareMessageCompleteBody strComputer, arrOut, strLog		' show All Output
	else
		PrepareMessageBody strComputer, arrOut, strLog
	end if
next

if strLog <> "" then 		' if its blank, don't send anything
	strLog = "Replication Report" & vbCrLf & vbCrLf & strLog
	SendLogMail EMAIL_TO, EMAIL_FROM, EMAIL_SUBJECT_LINE, strLog
end if 

Sub PrintReport(strComp, arrArray, filOut)
	For count = 1 to Ubound(arrOut) step 4
		
		arrtmp = split(arrArray(count),",")
		strContainer = arrtmp(0)
		strServ = "error"
		intPos = instr(arrArray(count+1), "via")-2
		
		if intPos>0 then 
			strServ = left(arrArray(count+1),intPos)
		else 
			strServ = "Error***********" & arrArray(count+1) & "***" & arrArray(count)
		end if
		
		if instr(arrArray(count+3),"successful") then 
			strStatus = "success"
		else
			strStatus = "fail"
		end if
		filOut.WriteLine strComputer & vbTab & strStatus & vbTab & strServ & vbTab & strContainer
	next
end sub

Sub PrepareMessageBody(strComp, arrArray, strLog)
	For count = 1 to Ubound(arrOut) step 4
		
		arrtmp = split(arrArray(count),",")
		strContainer = arrtmp(0)
		strServ = "error"
		intPos = instr(arrArray(count+1), "via")-2
		
		if intPos>0 then 
			strServ = left(arrArray(count+1),intPos)
		else 
			strServ = "Error***********" & arrArray(count+1) & "***" & arrArray(count)
		end if
		
		if instr(arrArray(count+3),"successful") then 
			strStatus = "success"
		else
			strStatus = "fail"
			strLog = strLog & vbCrLf & strComputer & vbTab & strStatus & vbTab & strServ & vbTab & strContainer
		end if
	next
end sub
Sub PrepareMessageCompleteBody(strComp, arrArray, strLog)
	For count = 1 to Ubound(arrOut) step 4
		
		arrtmp = split(arrArray(count),",")
		strContainer = arrtmp(0)
		strServ = "error"
		intPos = instr(arrArray(count+1), "via")-2
		
		if intPos>0 then 
			strServ = left(arrArray(count+1),intPos)
		else 
			strServ = "Error***********" & arrArray(count+1) & "***" & arrArray(count)
		end if
		
		if instr(arrArray(count+3),"successful") then 
			strStatus = "success"
		else
			strStatus = "fail"
		end if
		strLog = strLog & vbCrLf & strComputer & vbTab & strStatus & vbTab & strServ & vbTab & strContainer
	next
end sub
sub addArray(arrArray, strText)
	intTop = Ubound(arrArray)
	redim preserve arrArray(intTop+1)
	arrArray(intTop+1)=strText	
end sub

function ReadRepl (strCompname)
	Dim arrOutput()
	Redim arrOutput(0)
	set objScript = CreateObject("Wscript.Shell")
	' repadmin /showreps compname
	strCommand = "cmd /c repadmin /showreps " & strComputer
	'wscript.echo strCommand
	set objExec = objScript.Exec(strCommand)

	Do While Not objExec.StdOut.AtEndOfStream
		strHeader = objExec.StdOut.ReadLine()
		if (Instr(strHeader,"DC=") and (instr(strHeader,"Naming Context")=0)) then		' start at a section
			
			strInfos = objExec.StdOut.ReadLine()
			Do While len(strInfos)>5
				addArray arrOutput, trim(strHeader)	' add that section as a header
				'wscript.echo "strHeader: " & strHeader
				addArray arrOutput, trim(strInfos) '1 server rep pair
				'wscript.echo "#1: " & strInfos
				strInfos = objExec.StdOut.ReadLine()
				addArray arrOutput, trim(strInfos) '2 guid
				'wscript.echo "#2: " & strInfos
				strInfos = objExec.StdOut.ReadLine()
				addArray arrOutput, trim(strInfos) '3 success/fail
				'wscript.echo "#3: " & strInfos
				if instr(strInfos, "failed") then  ' burn 3 if its an error
					strInfos = objExec.StdOut.ReadLine()
					strInfos = objExec.StdOut.ReadLine()
					strInfos = objExec.StdOut.ReadLine()
				end if
				
				strInfos = objExec.StdOut.ReadLine()  ' next 1
			Loop
		end if 
	Loop
	'For x = Lbound(arrOutput) to Ubound(arrOutput)
	'	wscript.echo "ReadRepl " & arrOutput(x)
	'next
	
	ReadRepl = arrOutput
end function

function getControllers()

	Set objRootDSE = GetObject("LDAP://RootDSE")
	strADsPath = "LDAP://" & objRootDSE.Get("defaultNamingContext")
	Set objConn = CreateObject("ADODB.Connection")
	objConn.Open "Provider=ADsDSOObject;"
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConn

	objCommand.Properties("Page Size")=1000
	objCommand.Properties("Timeout")=30
	objCommand.Properties("Searchscope")=ADS_SCOPE_SUBTREE
	objCommand.Properties("Cache Results")=False

	strDomainControllerFilter = "(primaryGroupId=516)"
	objCommand.CommandText = "<" & strADsPath & ">;" & strDomainControllerFilter & ";sAMAccountName;subtree"

	Set objRecordSet=objCommand.Execute
	objRecordSet.MoveFirst
	Dim arrControllers()
	Redim arrControllers(0)
	Do Until objRecordSet.EOF
		strCompy = objRecordSet.Fields(0).Value
		'wscript.Echo left(strCompy, len(strCompy)-1)
		addArray arrControllers, left(strCompy, len(strCompy)-1)
		objRecordSet.MoveNext
	Loop
	objConn.Close 
	getControllers = arrControllers

end function

sub SendLogMail (strTo, strFrom, strSubject, strBody)

	Set cdoConfig = CreateObject("CDO.Configuration") 
	sch = "http://schemas.microsoft.com/cdo/configuration/" 
	cdoConfig.Fields.Item(sch & "sendusing") = 2 ' cdoSendUsingPort 
	cdoConfig.Fields.Item(sch & "smtpserver") = "smtp.adc.com" 
	cdoConfig.Fields.update 
	
	Set objMessage = CreateObject("CDO.Message")
	objMessage.Configuration = cdoConfig 
	objMessage.Subject = strSubject
	objMessage.Sender = strFrom
	objMessage.To = strTo
	objMessage.TextBody = strBody

	objMessage.Send

	Set objMessage = Nothing
	Set cdoConfig=nothing

end sub