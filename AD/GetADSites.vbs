
Dim arrSubnets()
Redim arrSubnets(0)

set objRootDSE = GetObject("LDAP://RootDSE")
set objSitesCont = GetObject("LDAP://cn=Subnets,cn=sites," & objRootDSE.Get("configurationNamingContext") )
objSitesCont.Filter = Array("subnet")
for each objSubnet in objSitesCont
   addArray arrSubnets, objSubnet.Get("name") 
next
arrTmpArray = CombSort(arrSubnets)

'for each subnet in arrSubnets
'	wscript.echo subnet
'next


sub addArray(arrArray, strText)
	intTop = Ubound(arrArray)
	redim preserve arrArray(intTop+1)
	arrArray(intTop+1)=strText			
end sub


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
