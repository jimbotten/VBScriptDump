arrDCs = getControllers
for each item in arrDCs
	wscript.echo item
next

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

sub addArray(arrArray, strText)
	intTop = Ubound(arrArray)
	redim preserve arrArray(intTop+1)
	arrArray(intTop+1)=strText	
end sub
