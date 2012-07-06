' get the command line arguments
set oArgs = WScript.Arguments
set oNamed = oArgs.Named

strName = oNamed("newname")
' cscript script.vbs /paramname:parameter

set wmiObj = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
changecomputername(wmiObj, strName)


Function ChangeComputerName(wmiObject, newName)
    For Each objComputer in wmiObject.InstancesOf("Win32_ComputerSystem")
        ChangeComputerName = objComputer.rename(newName)
    Next
End Function
