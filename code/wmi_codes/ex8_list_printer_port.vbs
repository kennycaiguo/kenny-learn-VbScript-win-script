strComputer = "."

set printerDict = CreateObject("Scripting.Dictionary")
set wmiObj = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

set items = wmiObj.ExecQuery("select * from Win32_Printer") 

For Each printer In items
    printerDict.Add printer.PortName ,printer.PortName
Next ' printer

set colPorts = wmiObj.ExecQuery("select * from Win32_TCPIPPrinterPort")

For Each port In colPorts
    If objDictionary.Exists(objPort.Name) Then
        strUsedPorts = strUsedPorts & _
            objDictionary.Item(objPort.Name) & VbCrLf
    Else
        strFreePorts = strFreePorts & objPort.Name & vbCrLf
    End If
Next ' port

Wscript.Echo "The following ports are in use: " & VbCrLf & strUsedPorts
Wscript.Echo "The following ports are available: " & VbCrLf & strFreePorts