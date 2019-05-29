On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_BIOS",,48)
For Each objItem in colItems

    Wscript.Echo "Naziv Izdelovalca: " & objItem.Manufacturer

    Wscript.Echo "Datum izdaje: " & objItem.ReleaseDate

Next
