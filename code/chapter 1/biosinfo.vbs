On Error Resume Next
Set wmiobj = GetObject("winmgmts:\\.\root\cimv2")

Set cols = wmiobj.ExecQuery("SELECT * FROM Win32_BIOS")

For Each bios In cols
   WScript.Echo "Name:" & bios.Name
   WScript.Echo "Serial Number:" & bios.SerialNumber
   WScript.Echo "Caption:" & bios.Caption 
   WScript.Echo "Status: " & bios.Status
   WScript.Echo "SMBIOSBIOSVersion:" & bios.SMBIOSBIOSVersion
Next 


