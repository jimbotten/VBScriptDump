Set objCommand = CreateObject("ADODB.Command")
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection

strBase = "<LDAP://server/DC=domain,DC=gov>"
strFilter = "(&(|(operatingSystem=Windows XP Professional)(operatingSystem=Windows 2000 Professional))(objectClass=computer))"
strAttributes = "Name,dnshostname,operatingSystem,operatingSystemVersion"
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
objCommand.CommandText = strQuery
objCommand.Properties("Page Size") = 100
objCommand.Properties("Timeout") = 30
objCommand.Properties("Cache Results") = False
Set objRecordSet = objCommand.Execute

Do Until objRecordSet.EOF
  strName = objRecordSet.Fields("Name").Value
  strDNS = objRecordSet.Fields("dnshostname").value
  strOS = objRecordSet.Fields("operatingSystem").value 
  strOSVer = objRecordSet.Fields("operatingSystemVersion").value
  'strIP = pingComp(strName)
  
  wscript.echo strName _
	& "|" & strDNS _
	& "|" & strOS _
	& "|" & strOSVer 
	'& "|" & strIP 
	
  objRecordSet.MoveNext
Loop
 
objConnection.Close

function pingComp(strComputer)
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colPings = objWMIService.ExecQuery ("Select * From Win32_PingStatus where Address = '" & strComputer & "'")

    For Each objStatus in colPings
        If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then 
            pingComp = 0
        Else
            pingComp = objStatus.ProtocolAddress
        End If
    Next
End function
