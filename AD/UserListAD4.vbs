Set objCommand = CreateObject("ADODB.Command")
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection

strBase = "<LDAP://ou=users,dc=domain,dc=com>"
strFilter = "(&(objectCategory=person)(objectClass=user))"
strAttributes = "sAMAccountName,cn,objectSID,description,AssocNTAccount"
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
objCommand.CommandText = strQuery
objCommand.Properties("Page Size") = 100
objCommand.Properties("Timeout") = 30
objCommand.Properties("Cache Results") = False
Set objRecordSet = objCommand.Execute

Do Until objRecordSet.EOF
  strName = objRecordSet.Fields("sAMAccountName").Value
  strCN = objRecordSet.Fields("cn").value
  strSID = objRecordSet.Fields("objectSID").value
  strDesc = objRecordSet.Fields("description").value
  strAssoc = objRecordSet.Fields("AssocNTAccount").value
  wscript.echo "CORP\" & strName _
	& "|" & HexStrToDecStr(OctetToHexStr(strSID)) _
	& "|" & strDesc _
	& "|" & HexStrToDecStr(OctetToHexStr(strAssoc))
  objRecordSet.MoveNext
Loop
 
objConnection.Close

Function OctetToHexStr(arrbytOctet)
  '  Function to convert OctetString (byte array) to Hex string.
  Dim k
  OctetToHexStr = ""
  For k = 1 To Lenb(arrbytOctet)
    OctetToHexStr = OctetToHexStr & Right("0" & Hex(Ascb(Midb(arrbytOctet, k, 1))), 2)
  Next
End Function

Function HexStrToDecStr(strSid)
  ' Function to convert hex Sid to decimal (SDDL) Sid.
  Dim arrbytSid, lngTemp, j

  ReDim arrbytSid(Len(strSid)/2 - 1)
  For j = 0 To UBound(arrbytSid)
    arrbytSid(j) = CInt("&H" & Mid(strSid, 2*j + 1, 2))
  Next

  HexStrToDecStr = "S-" & arrbytSid(0) & "-" & arrbytSid(1) & "-" & arrbytSid(8)

  lngTemp = arrbytSid(15)
  lngTemp = lngTemp * 256 + arrbytSid(14)
  lngTemp = lngTemp * 256 + arrbytSid(13)
  lngTemp = lngTemp * 256 + arrbytSid(12)

  HexStrToDecStr = HexStrToDecStr & "-" & CStr(lngTemp)
 
  lngTemp = arrbytSid(19)
  lngTemp = lngTemp * 256 + arrbytSid(18)
  lngTemp = lngTemp * 256 + arrbytSid(17)
  lngTemp = lngTemp * 256 + arrbytSid(16)

  HexStrToDecStr = HexStrToDecStr & "-" & CStr(lngTemp)

  lngTemp = arrbytSid(23)
  lngTemp = lngTemp * 256 + arrbytSid(22)
  lngTemp = lngTemp * 256 + arrbytSid(21)
  lngTemp = lngTemp * 256 + arrbytSid(20)
 
  HexStrToDecStr = HexStrToDecStr & "-" & CStr(lngTemp)
  
  lngTemp = arrbytSid(25)
  lngTemp = lngTemp * 256 + arrbytSid(24)
 
  HexStrToDecStr = HexStrToDecStr & "-" & CStr(lngTemp)

End Function