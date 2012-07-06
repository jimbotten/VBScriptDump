strComputer = "."
Set System = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colShares = System.ExecQuery("SELECT * FROM Win32_Share")
For Each objShare in colShares
	wscript.echo "---------------------------" & vbcrlf & objShare
	EnumerateSharePermission(objShare)
Next

Function EnumerateSharePermission(strShareName)

On Error Resume Next
' The Win32_LogicalShareSecuritySetting instance with
' the name = to WMILogs$ is specified 
Set wmiFileSecSetting = GetObject( "winmgmts:Win32_LogicalShareSecuritySetting.Name='"& strShareName &"$'")

RetVal = wmiFileSecSetting.GetSecurityDescriptor(wmiSecurityDescriptor)
If Err <> 0 Then
    WScript.Echo "GetSecurityDescriptor failed" & VBCRLF & Err.Number & VBCRLF & Err.Description
    WScript.Quit
Else
    WScript.Echo "GetSecurityDescriptor succeeded"
End If

' Retrieve the DACL array of Win32_ACE objects.
DACL = wmiSecurityDescriptor.DACL

For each wmiAce in DACL

    wscript.echo "Access Mask: "     & wmiAce.AccessMask
    wscript.echo "ACE Type: "        & wmiAce.AceType

' Get Win32_Trustee object from ACE 
       Set Trustee = wmiAce.Trustee
    wscript.echo "Trustee Domain: "  & Trustee.Domain
    wscript.echo "Trustee Name: "    & Trustee.Name

' Get SID as array from Trustee
    SID = Trustee.SID 
    strsid = join(SID, ",") 
    wscript.echo "Trustee SID: {" & strsid & "}"
        
Next
End Function