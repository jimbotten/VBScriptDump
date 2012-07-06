Set WshNetwork = WScript.CreateObject("WScript.Network")
Set oPrinters = WshNetwork.EnumPrinterConnections
         
Wscript.Echo "*** PRINTER UPDATES ***"
		 
For i = 0 to oPrinters.Count - 1 Step 2		' cycle through all the printers on the system
	intPlace=0
	strCurrPrinter = oPrinters.Item(i+1)
	WScript.Echo "Found " & strCurrPrinter
Next

set oPrinters = nothing
set WshNetwork = nothing