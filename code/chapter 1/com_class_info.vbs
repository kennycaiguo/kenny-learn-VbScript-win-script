
Set wmi = GetObject("winmgmts:\\.\root\cimv2")
Set cols = wmi.ExecQuery("SELECT * FROM Win32_COMClass")

For Each ctr In cols
  WScript.Echo "Name: " & ctr.Name
  WScript.Echo "Description:" & ctr.Description

Next


