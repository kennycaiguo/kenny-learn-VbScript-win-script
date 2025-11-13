On Error Resume Next
strcomputer = "."

Set wmiObj = GetObject _
  ("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strcomputer & "\root\cimv2")

Set colItems = wmiObj.ExecQuery("select * from Win32_ApplicationService")

For Each item In colItems
    WScript.Echo "Name:" & item.Name
    WScript.Echo "Start Mode:" & item.StartMode
    WScript.Echo
Next ' item