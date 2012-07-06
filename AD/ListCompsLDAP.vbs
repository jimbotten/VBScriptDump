Const ADS_UF_PASSWD_CANT_CHANGE = &H40
Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000

Dim strFilePath, objFSO, objFile, objConnection, objCommand
Dim objRootDSE, strDNSDomain, strFilter, strQuery, objRecordSet
Dim strDN, objShell, lngBiasKey, lngBias, blnPwdExpire
Dim lngDate, objDate, dtmPwdLastSet, lngFlag, k

Set objShell = CreateObject("Wscript.Shell")

strDN = InputBox("Enter the distinguished name of a container.", "Container", "OU=compsDC=domain,DC=gov")
strAttributes = InputBox("Enter the attributes to return.", "Attributes", "name,distinguishedname")
strFilterEntry = InputBox("Enter the type of object to filter by. (user, computer)", "Filter", "computer")
strDelim = InputBox("What is the field delimiter?  how about a pipe?", "Delimiter", "|")

arrAttributes = split(strAttributes, ",")
strFilter = "(objectCategory=" & strFilterEntry & ")"

' Use ADO to search the domain for all users.
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOOBject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

' Determine the DNS domain from the RootDSE object.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("DefaultNamingContext")

strQuery = "<LDAP://" & strDN & ">;" & strFilter _
  & ";" & strAttributes & ";subtree"

objCommand.CommandText = strQuery
objCommand.Properties("Page Size") = 100
objCommand.Properties("Timeout") = 30
objCommand.Properties("Cache Results") = False

' Enumerate all users. Write each user's Distinguished Name,
' whether they are allowed to change their password, and when
' they last changed their password to the file.
wscript.echo strQuery

Set objRecordSet = objCommand.Execute

Do Until objRecordSet.EOF
  strout = ""

  intCount = objRecordset.Fields.Count
  For x = 0 to intCount-1
   strout = strout & objRecordset.Fields(x) & strDelim
   'msgbox strout
  Next 

  wscript.echo strout
  objRecordSet.MoveNext
Loop