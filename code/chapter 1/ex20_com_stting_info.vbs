Set wmi = GetObject("winmgmts:\\.\root\cimv2")

Set cols = wmi.ExecQuery("SELECT * FROM Win32_COMSetting")

For Each setting In cols
   WScript.Echo "Caption:" & setting.Caption
   WScript.Echo "Description:" & setting.Description
   WScript.Echo "SettingID:" & setting.SettingID
Next
