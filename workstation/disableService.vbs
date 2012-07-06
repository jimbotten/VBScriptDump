strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colServiceList = objWMIService.ExecQuery _
    ("Select * from Win32_Service where Name = 'lanmanserver'")

For Each objService in colServiceList
    If objService.State = "Running" Then
        objService.StopService()
        Wscript.Sleep 5000
    End If
    errReturnCode = objService.ChangeStartMode("Disabled")   
Next