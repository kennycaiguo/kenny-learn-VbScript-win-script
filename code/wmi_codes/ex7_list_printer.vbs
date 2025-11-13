strComputer = "."

set wmiService = GetObject _ 
  ("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

set colItems = wmiService.ExecQuery("select * from Win32_Printer")  

For Each item In colItems
  WScript.Echo "Name:" & item.Name
  WScript.Echo "Location:" & item.Location
  WScript.Echo "Default:" & item.Default
Next 'item