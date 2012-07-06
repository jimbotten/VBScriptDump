Const FROMsERVER = "oldPrintServer"
Const TOsERVER =  "newPrintServer"


Set WshNetwork = WScript.CreateObject("WScript.Network")
Set oPrinters = WshNetwork.EnumPrinterConnections
         
Wscript.Echo "*** PRINTER UPDATES ***"
		 
For i = 0 to oPrinters.Count - 1 Step 2
	WScript.Echo "Found " & oPrinters.Item(i+1) & vbTab & "Printer: " & oPrinters.Item(i)

	intPlace = instr(oPrinters.Item(i+1), "\\" & FROMsERVER & "\")
	if intPlace then
		strPrinter = "\\" & TOsERVER & "\" & right(oPrinters.Item(i+1), len(oPrinters.Item(i+1))-11)
		'Remove old
		WshNetwork.RemovePrinterConnection oPrinters.Item(i+1), true, true
		wscript.echo vbTab & "----> Adding " & strPrinter
		' Add New
		WshNetwork.AddWindowsPrinterConnection strPrinter
		end if
Next

set oPrinters = nothing
set WshNetwork = nothing