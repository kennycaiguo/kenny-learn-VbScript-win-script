Set wmi = GetObject("winmgmts:\\.\root\cimv2")

Set cols = wmi.ExecQuery("SELECT * FROM Win32_ComputerSystemProcessor")

For Each cpu In cols
 WScript.Echo "GroupComponent:" & cpu.GroupComponent
 WScript.Echo "PartComponent:" & cpu.PartComponent
Next