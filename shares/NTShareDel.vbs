'	Script 3/16/05 by Jim Montgomery
'	This script will {try to}connect to the machine name and remove the specified share name

	strComputer= InputBox("What host to remove a share from?")
	strShareName = InputBox("What is the sharename?")

	Err.Clear
	Set System = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	if (Err.Number <> 0) then
		msgbox("Error: " & Err.Description)
	else
		Set colShares = System.ExecQuery("SELECT * FROM Win32_SHare where Name = '" & strShareName & "'")
			if (Err.Number <> 0) then
				msgbox("Error: " & Err.Description)
			else 
				For Each objShare in colShares
					objShare.Delete
				Next
			end if		
	end if 

set colShares= nothing
set System = nothing

msgbox("Script Finished, share removed.")