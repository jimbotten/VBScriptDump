CONST TODomain = "domain"
strAdminUser = "domain\admin"
strAdminPassword = "pass"

strInputFile = inputbox("What is the path to the input file?","Input","C:\comps.txt")
Set objExcel = WScript.CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add
objExcel.ActiveSheet.Name = "Migrate"
objExcel.ActiveSheet.Range("A1").Activate
objExcel.ActiveCell.Value = "Computer"		'col header 1
objExcel.ActiveCell.Offset(0,1).Value = "Pingable"
objExcel.ActiveCell.Offset(0,2).Value = "Name"
objExcel.ActiveCell.Offset(0,3).Value = "Domain"
objExcel.ActiveCell.Offset(0,4).Value = "Reboot"

objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down

' ******************************

'On Error Resume Next
set objFSO = CreateObject("Scripting.FileSystemObject")
set objFile = objFSO.OpenTextFile(strInputFile)
Do Until objFile.AtEndOfStream
	cancel = 0
	named = 0
	domained = 0
	arrLine = split(objFile.ReadLine,",")
	strComputer= arrLine(0)
	strNewComputer = arrLine(1)
	
	objExcel.ActiveCell.Value = strComputer
	'check ping
	if pingComp(strComputer) = 0 then
		objExcel.ActiveCell.Offset(0,1).Value = "No connection"
		cancel=1
	else
		objExcel.ActiveCell.Offset(0,1).Value = "Pingable"
	end if 
	
	set wmiObject = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	
	'checkname
	if strComputer <> strNewComputer and cancel = 0 then
		ChangeComputerName wmiObject, strNewComputer
		named = 1
		objExcel.ActiveCell.Offset(0,2).Value = "Changed to " & strNewComputer
	else
		objExcel.ActiveCell.Offset(0,2).Value = "No Change"
	end if
	
	'changedomain
	strCompDom = GetComputerDomain(wmiObject)
	if strCompDom <> TODomain and cancel = 0 then
		' add "error"
		objExcel.ActiveCell.Offset(0,3).Value = "error" 
		' add more if possible.  seperate line.
		objExcel.ActiveCell.Offset(0,3).Value = objExcel.ActiveCell.Offset(0,3).Value & fncErrorMessage(joinDomain(wmiObject, TODomain, strAdminUser, strAdminPassword))
		domained = 1
	else
		objExcel.ActiveCell.Offset(0,3).Value = "No Change"
	end if

	'reboot
	if named = 1 or domained = 1 then
		reboot(wmiObject)
		objExcel.ActiveCell.Offset(0,4).Value = "Rebooted"
	else
		objExcel.ActiveCell.Offset(0,4).Value = "Nope"
	end if
	
	objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down
	objFile.Movenext
Loop

set objFile = nothing
Set objFSO = nothing
set objExcel=nothing
msgbox("Script Finished")

' **********************
function pingComp(strComputer)
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colPings = objWMIService.ExecQuery ("Select * From Win32_PingStatus where Address = '" & strComputer & "'")

    For Each objStatus in colPings
        If IsNull(objStatus.StatusCode) _
            or objStatus.StatusCode<>0 Then 
            pingComp = 0
        Else
            pingComp = 1
        End If
    Next
End function

Function ChangeComputerName(wmiObject, newName)
    For Each objComputer in wmiObject.InstancesOf("Win32_ComputerSystem")
        ChangeComputerName = objComputer.rename(newName)
    Next
End Function

function GetComputerDomain(wmiObject)
	' This returns either a 0 if not in a domain, or the name of the domain its part of
	'on error resume next
	
  	wql = "select DomainRole from Win32_ComputerSystem"
	set results = wmiObject.execquery(wql)
	for each obj in results
		strDomainRole = obj.DomainRole
	next
'msgbox "domainRole for this computersystem: " & strDomainRole
	if strDomainRole <> 0 and strDomainRole <> 2 then
	    wql = "select Domain from Win32_ComputerSystem"
	    set results = wmiObject.execquery(wql)
	    for each obj in results
    		GetComputerDomain = obj.Domain
			'msgbox obj.domain
	    next
    else
        GetComputerDomain = 0
	end if
	
'Value Meaning 
'0x0 Standalone Workstation 
'0x1 Member Workstation 
'0x2 Standalone Server 
'0x3 Member Server 
'0x4 Backup Domain Controller 
'0x5 Primary Domain Controller 

end function 

Function joinDomain(wmiObject, strDomain, strUser, strPass)
	Const JOIN_DOMAIN             = 1
    Const ACCT_CREATE             = 2
    Const ACCT_DELETE             = 4
    Const WIN9X_UPGRADE           = 16
    Const DOMAIN_JOIN_IF_JOINED   = 32
    Const JOIN_UNSECURE           = 64
    Const MACHINE_PASSWORD_PASSED = 128
    Const DEFERRED_SPN_SET        = 256
    Const INSTALL_INVOCATION      = 262144
    For Each client in wmiObject.InstancesOf("Win32_ComputerSystem")
        joinDomain = Client.JoinDomainOrWorkGroup(strDomain, strPass, strUser, NULL, JOIN_DOMAIN)
    next
end function

Function fncErrorMessage(varErrorNumber)
        Select Case varErrorNumber
            Case 0 strErrorDescription = "Success"
			Case 5 strErrorDescription = "Access is denied"
            Case 87 strErrorDescription = "The parameter is incorrect"
            Case 110 strErrorDescription = "The system cannot open the specified object"
            Case 1323 strErrorDescription = "Unable to update the password"
            Case 1326 strErrorDescription = "Logon failure: unknown username or bad password"
            Case 1355 strErrorDescription = "The specified domain either does not exist or could not be contacted"
            Case 2224 strErrorDescription = "The account already exists"
            Case 2691 strErrorDescription = "The machine is already joined to the domain"
            Case 2692 strErrorDescription = "The machine is not currently joined to a domain"
        End Select
    fncErrorMessage = "Error: " & varErrorNumber & ". " & strErrorDescription & "."
End Function

sub reboot(objWMI)
			'kill nav
			Set colServiceList = objWMI.ExecQuery("SELECT * FROM Win32_Service WHERE Name='nalntservice'")
			For Each objService in colServiceList
				errReturn = objService.StopService()
			Next
			For Each objComputer in wmiObject.InstancesOf("Win32_OperatingSystem")
				objComputer.reboot()
			Next
end sub