strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_NetworkProtocol")

For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Connectionless Service: " & objItem.ConnectionlessService
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Guarantees Delivery: " & objItem.GuaranteesDelivery
    Wscript.Echo "Guarantees Sequencing: " & objItem.GuaranteesSequencing
    strInstallDate = WMIDateStringToDate(objItem.InstallDate)
    Wscript.Echo "Install Date: " & strInstallDate
    Wscript.Echo "Maximum Address Size: " & objItem.MaximumAddressSize
    Wscript.Echo "Maximum Message Size: " & objItem.MaximumMessageSize
    Wscript.Echo "Message Oriented: " & objItem.MessageOriented
    Wscript.Echo "Minimum Address Size: " & objItem.MinimumAddressSize
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Pseudo Stream Oriented: " & objItem.PseudoStreamOriented
    Wscript.Echo "Supports Broadcasting: " & objItem.SupportsBroadcasting
    Wscript.Echo "Supports Connect Data: " & objItem.SupportsConnectData
    Wscript.Echo "Supports Disconnect Data: " & objItem.SupportsDisconnectData
    Wscript.Echo "Supports Encryption: " & objItem.SupportsEncryption
    Wscript.Echo "Supports Expedited Data: " & objItem.SupportsExpeditedData
    Wscript.Echo "Supports Fragmentation: " & objItem.SupportsFragmentation
    Wscript.Echo "Supports Graceful Closing: " & _
        objItem.SupportsGracefulClosing
    Wscript.Echo "Supports Guaranteed Bandwidth: " & _
        objItem.SupportsGuaranteedBandwidth
    Wscript.Echo "Supports Multicasting: " & objItem.SupportsMulticasting
    Wscript.Echo "Supports Quality of Service: " & _
        objItem.SupportsQualityofService
Next
 
Function WMIDateStringToDate(dtmDate)
    WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
        Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
            & " " & Mid (dtmDate, 9, 2) & ":" & _
                Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function