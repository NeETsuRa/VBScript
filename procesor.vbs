On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor",,48)
For Each objItem in colItems
    	Wscript.Echo "Opis: " & objItem.Description
    	Wscript.Echo "Proizvajalec: " & objItem.Manufacturer
    	Wscript.Echo "tip podnožja: " & objItem.SocketDesignation
Next
