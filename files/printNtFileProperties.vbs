Set objExcel = WScript.CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add
objExcel.ActiveSheet.Name = "Files"
objExcel.ActiveSheet.Range("A1").Activate
objExcel.ActiveCell.Value = "Path"		'col header 1
objExcel.ActiveCell.Offset(0,1).Value = "Name"
objExcel.ActiveCell.Offset(0,2).Value = "CreationDate"
objExcel.ActiveCell.Offset(0,3).Value = "LastAccessed"
objExcel.ActiveCell.Offset(0,4).Value = "LastModified"
objExcel.ActiveCell.Offset(0,5).Value = "Attributes"

objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down
'-------------------------

strComputer = InputBox("Computer",,".")
strFolder = InputBox("Look at which folder's permissions",,"c:\script\test")
Set FSO = CreateObject("Scripting.FileSystemObject")


ShowSubfolders FSO.GetFolder(strFolder)

Sub ShowSubFolders(Folder)
Dim intAtty
 For Each Subfolder in Folder.SubFolders

Set colFiles = Subfolder.Files

	objExcel.ActiveCell.Value = Subfolder.Path
	objExcel.ActiveCell.Offset(0,1).Value = 	Subfolder.Name
	objExcel.ActiveCell.Offset(0,2).Value = 	Subfolder.DateCreated
	objExcel.ActiveCell.Offset(0,3).Value = 	Subfolder.DateLastAccessed
	objExcel.ActiveCell.Offset(0,4).Value = 	Subfolder.DateLastModified
'	objExcel.ActiveCell.Offset(0,5).Value = 	ShowAtt(Subfolder.Attributes)
	objExcel.ActiveCell.Offset(0,5).Value = 	Subfolder.Attributes
	objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down

For Each objFile in colFiles
	objExcel.ActiveCell.Value = objFile.Path
	objExcel.ActiveCell.Offset(0,1).Value = 	objFile.Name
	objExcel.ActiveCell.Offset(0,2).Value = 	objFile.DateCreated
	objExcel.ActiveCell.Offset(0,3).Value = 	objFile.DateLastAccessed
	objExcel.ActiveCell.Offset(0,4).Value = 	objFile.DateLastModified
'	objExcel.ActiveCell.Offset(0,5).Value = 	ShowAtt(IntAtty)
	objExcel.ActiveCell.Offset(0,5).Value = 	objFile.Attributes
	objExcel.ActiveCell.Offset(1,0).Activate	'move 1 down
Next

 ShowSubFolders Subfolder
 Next
End Sub

Sub ShowAtt(Att)

'Constant 		Value 	Description
'Normal 		0 	Normal file. No attributes are set.
'ReadOnly 	1 	Read-only file. Attribute is read/write.
'Hidden 		2 	Hidden file. Attribute is read/write.
'System 		4 	System file. Attribute is read/write.
'Volume 		8 	Disk drive volume label. Attribute is read-only.
'Directory 		16 	Folder or directory. Attribute is read-only.
'Archive 		32 	File has changed since last backup. Attribute is read/write.
'Alias 		1024 	Link or shortcut. Attribute is read-only.
'Compressed 	2048 	Compressed file. Attribute is read-only.

msgbox("got this far")

if (1) then 
	strAtt = "Ro"  
else 
	strAtt = "Rw" 
end if

ShowAtt = strAtt

End Sub