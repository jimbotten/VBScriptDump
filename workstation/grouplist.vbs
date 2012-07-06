CRLF=CHR(13)+CHR(10)

strDC = "computer or DC"

Set DomainObj = GetObject("Winnt://" & strDC)
if Err.Number <0 then
wscript.echo "Failed to connect to " & strDC
wscript.quit
end if
DomainObj.Filter = Array("group")

For Each GroupObj In DomainObj
	wscript.echo GroupObj.Name
next