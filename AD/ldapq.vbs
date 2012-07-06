Const ADS_UF_PASSWD_CANT_CHANGE = &H40
Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000

Dim strFilePath, objFSO, objFile, objConnection, objCommand
Dim objRootDSE, strDNSDomain, strFilter, strQuery, objRecordSet
Dim strDN, objShell, lngBiasKey, lngBias, blnPwdExpire
Dim lngDate, objDate, dtmPwdLastSet, lngFlag, k

Set objShell = CreateObject("Wscript.Shell")

' Use ADO to search the domain for all users.
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOOBject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 100
objCommand.Properties("Timeout") = 30
objCommand.Properties("Cache Results") = False

' Determine the DNS domain from the RootDSE object.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("DefaultNamingContext")

' **********Retrieve Users &***************
strFilter = "(objectCategory=user)"
strQuery = "<LDAP://" & strDNSDomain & ">;" & strFilter & ";cn;subtree"
objCommand.CommandText = strQuery
wscript.echo "Searching Users..."
Set objRecordSet = objCommand.Execute
wscript.echo objRecordSet.recordcount

' **********Retrieve Computers (servers) ***************
strFilter = "(objectCategory=computer)"
strQuery = "<LDAP://" & strDNSDomain & ">;" & strFilter & ";cn,operatingSystem;subtree"
objCommand.CommandText = strQuery
wscript.echo "Searching Servers..."
Set objRecordSet = objCommand.Execute
wscript.echo objRecordSet.recordcount

Do Until objRecordSet.EOF
  strDN = objRecordSet.Fields("operatingsystem")
  wscript.echo strDN
  objRecordSet.MoveNext
Loop

