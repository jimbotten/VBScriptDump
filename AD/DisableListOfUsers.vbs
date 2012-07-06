set oArgs = WScript.Arguments
set oNamed = oArgs.Named
strDisable = oNamed("d")

if oNamed.count = 0 then
	wscript.echo "TEST RUN.  To run this command to disable, use parameter:" & vbCrLf & "/d:1"
end if

' This is the disable bit
Const ADS_UF_ACCOUNTDISABLE = 2

' Set up the ADO search objects
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Cache Results") = False
objCommand.Properties("Page Size") = 1000

strInputFile = inputBox("What is the file name of the list of users?","Disable Users","c:\users.txt")
set objFSO = CreateObject("Scripting.FileSystemObject")
set objFile = objFSO.OpenTextFile(strInputFile)

' Determine DNS domain name.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")
strAttributes = "AdsPath"
strBase = "<LDAP://" & strDNSDomain & ">"
'strBase = "<LDAP://dc=aa,dc=ad,dc=epa,dc=gov>"
'wscript.echo "Searching " & strBase

Do Until objFile.AtEndOfStream
	strUsername= objFile.ReadLine
		
	strFilter = "(&(&(objectCategory=person)(objectClass=user))(sAMAccountName=" & strUsername & "))"

	strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"	
	objCommand.CommandText = strQuery
	
	Set objRecordSet = objCommand.Execute
	if objRecordset.EOF then
		wscript.echo strUsername & vbTab & "Not Found"
	end if
	if not objRecordset.EOF then
	objRecordSet.MoveFirst
	Do Until objRecordSet.EOF
		strAdsPath = objRecordset.Fields("AdsPath").Value
				
		Set objUser = GetObject(strAdsPath)
		
		intlastlogontime = 0		
		
		on error resume next
		Set objLastLogon = objUser.Get("lastLogonTimestamp")
		
		intLastLogonTime = objLastLogon.HighPart * (2^32) + objLastLogon.LowPart 
		intLastLogonTime = intLastLogonTime / (60 * 10000000)
		intLastLogonTime = intLastLogonTime / 1440		
		'objUser.AccountDisabled = True
		
		if strDisable <> "" then
			intUAC = objUser.Get("userAccountControl")
			objUser.Put "userAccountControl", intUAC OR ADS_UF_ACCOUNTDISABLE
			objUser.SetInfo	
			wscript.echo "DISABLED" & vbTab & strUsername & vbTab & (intLastLogonTime + #1/1/1601#) & vbTab & strAdsPath
		else 
			wscript.echo strUsername & vbTab & (intLastLogonTime + #1/1/1601#) & vbTab & strAdsPath
		end if
		set objUser = nothing
		Set objLastLogon = nothing
		
		objRecordSet.MoveNext
	Loop
	end if 
Loop

set objFile = nothing
set objFSO = nothing

msgbox ("Script Complete")



function LastLogon(objDate)
                lngHigh = objDate.HighPart
                lngLow = objDate.LowPart
                If (lngLow < 0) Then
                    lngHigh = lngHigh + 1
                End If
                If (lngHigh = 0) And (lngLow = 0 ) Then
                    dtmDate = #1/1/1601#
                Else
                    dtmDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
                        + lngLow)/600000000 - lngBias)/1440
                End If
	LastLogon = dtmDate
end function