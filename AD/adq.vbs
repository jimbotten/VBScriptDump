Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
Set adoCommand.ActiveConnection = adoConnection

' Search entire Active Directory domain.
Set objRootDSE = GetObject("LDAP://RootDSE")

strDNSDomain = objRootDSE.Get("defaultNamingContext")
strBase = "<LDAP://" & strDNSDomain & ">"

' Filter on user objects.
'strFilter = "(&(objectCategory=person)(objectClass=user)(!(mail=*))(employeeid=*))"
strFilter = "(&(objectCategory=person)(objectClass=user)(!(mail=*))(employeeid=*))"

' Comma delimited list of attribute values to retrieve.
'strAttributes = "samaccountname,cn,mail,givenname,surname,employeeid"
strAttributes = "samaccountname,cn,givenname,sn,employeeid"

' Construct the LDAP syntax query.
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 30
adoCommand.Properties("Cache Results") = False

' Run the query.
Set adoRecordset = adoCommand.Execute

' Enumerate the resulting recordset.
Do Until adoRecordset.EOF
	if isArray(adoRecordset.fields(0)) then	 
	  theArray = adoRecordset.fields(0)
	   output=output &  theArray(0)
	else
	  output=output & adoRecordset.fields(0)
	end if
	output=output &   ","
	if isArray(adoRecordset.fields(1)) then	 
	  theArray = adoRecordset.fields(1)
	   output=output & theArray(0)
	else
	  output=output & adoRecordset.fields(1)
	end if
	output=output & ","
	if isArray(adoRecordset.fields(2)) then	 
	  theArray = adoRecordset.fields(2)
	   output=output & theArray(0)
	else
	  output=output & adoRecordset.fields(2)
	end if
	output=output & ","
	if isArray(adoRecordset.fields(3)) then	 
	  theArray = adoRecordset.fields(3)
	   output=output & theArray(0)
	else
	  output=output & adoRecordset.fields(3)
	end if
	output=output & ","
	if isArray(adoRecordset.fields(4)) then	 
	  theArray = adoRecordset.fields(4)
	   output=output & theArray(0)
	else
	  output=output & adoRecordset.fields(4)
	end if
    adoRecordset.MoveNext
	wscript.echo output
	output = ""
Loop