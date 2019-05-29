On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_DiskDrive",,48)
st = 0
For Each objItem in colItems
    st=st + 1
    Wscript.Echo "Disk: " &st
    Wscript.Echo "Vmesnik: " & objItem.InterfaceType
    Wscript.Echo "Model: " & objItem.Model
    Wscript.Echo "Velikost: " & objItem.Size
    Wscript.Echo "--------"
Next
Wscript.Echo "Stevilo diskov: " &st