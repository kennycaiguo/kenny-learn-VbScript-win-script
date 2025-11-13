Set wmi = GetObject("winmgmts:\\.\root\cimv2")

Set cols = wmi.ExecQuery("SELECT * FROM Win32_DCOMApplication")

For Each dcom In cols
 WScript.Echo "APP id:" & dcom.AppID
 WScript.Echo "Caption: " & dcom.Caption
 WScript.Echo "Description:" & dcom.Description
 WScript.Echo "Name:" & dcom.Name
 WScript.Echo "Status:" & dcom.Status
 
 WScript.Echo "=============================="
Next