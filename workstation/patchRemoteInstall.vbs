'*  remote Patch Installation
'* 11/11/04  Jim Montgomery

ip = inputbox("Remote Hostname or IP?")
patchFileName = inputbox("Patch file name on your C drive?")

Const OverwriteExisting = True
set ObjFSO = CreateObject("Scripting.FileSystemObject")
objFSO.CopyFile "C:\"& patchFileName, "\\"& ip & "\c$\" & patchFileName, OverwriteExisting

set osvcRemote = GetObject("winmgmts:\\" & ip & "\root\cimv2")

set oprocess = osvcRemote.Get("win32_process")

'* quiet mode, no reboot when done.

ret = oprocess.create("c:\\" & patchFileName & " -quiet -norestart")

if (ret <> 0) then
	wscript.echo "Failed to start process on " & ip & ": " & ret
else
	Msgbox( "Patch installing.")
end if


msgbox ("Script Complete")


set oprocess = nothing
set osvcRemote = nothing




