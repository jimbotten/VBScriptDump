' This Script wil Iterate through a folder and remove the access for a user/group (trustee)
' from the DACL of a file or folder.
' Developed by Wasi Rizvi

Const tte = "Everyone"

Set objFSO = CreateObject("scripting.filesystemobject")
CheckFolder (objFSO.getfolder("c:\script\test"))


Sub CheckFolder( objCurrentFolder )

Dim objNewFolder
Dim objFile
Dim sDirPath
Dim oFSO,oFile

wscript.echo objCurrentFolder.path & "********"
GetFileDetails objCurrentFolder.path

For Each objFile In objCurrentFolder.Files
wscript.echo objFile.path & "********"
GetFileDetails objFile.path
Next

'Recurse through all of the folders
For Each objNewFolder In objCurrentFolder.subFolders
wscript.echo objNewFolder.path & "++++++++++++++++++++++++++++folder+++++++++"
GetFileDetails objNewFolder
CheckFolder objNewFolder
Next


End Sub

Sub GetFileDetails ( strFile )

WScript.Echo "File =>" & strFile
Set wmiFileSecuritySetting = _
GetObject("winmgmts:Win32_LogicalFileSecuritySetting='" & strFile & "'")

' Fetch the security descriptor, and store it in wmiSD.
retval = wmiFileSecuritySetting.GetSecurityDescriptor(wmiSD)
' Retrieve the information from the security descriptor.
Set wmiDacl = wmiSD.Properties_.Item("Dacl")

WScript.Echo strFile & " DACL" & vbCrLf & String(Len(strFile) + 5, "=")
For i = 0 To UBound(wmiDacl.Value)
WScript.Echo " Trustee: " & _
wmiDacl.Value(i).Properties_.Item("Trustee").Value.Properties_.Item("Name")
'WScript.Echo " AccessMask: " & wmiDacl.Value(i).Properties_.Item("AccessMask").Value
'WScript.Echo " AceFlags: " & wmiDacl.Value(i).Properties_.Item("AceFlags").Value


WScript.Echo "Starting to Delete ACE for " & tte & " from " & " file" & strFile



if ( wmiDacl.Value(i).Properties_.Item("Trustee").Value.Properties_.Item("Name") = tte ) then
WScript.Echo "Gotcha%%%%%%%%%%%%%%%%%%%%%%Gotcha%%%%%%%%%%%%%%%%%%%%%%Gotcha"


'**************************************************************************************Help Required at this point
'at this point what i am trying to do is to remove the tte from the DACL (wmiDacl.Value(i))
'But I am unable to find any function/method to do that that is supported by the
'object => wmiDacl
'
'Can you let me know as to how can it be done.
'Thanks a lot.
'**************************************************************************************

end if

WScript.Echo "Successfully Deleted ACE for " & tte & " from " & " file" & strFile

Next
End Sub
