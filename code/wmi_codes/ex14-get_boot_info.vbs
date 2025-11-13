'获取操作系统的启动配置信息
Set wmi = GetObject("winmgmts:\\.\root\cimv2")

Set cols = wmi.ExecQuery("SELECT * FROM Win32_BootConfiguration")

For Each conf In cols
  WScript.Echo "Boot Dir:" & conf.BootDirectory
  WScript.Echo "Caption:" & conf.Caption
  WScript.Echo "Configuration Path:" & conf.ConfigurationPath
  WScript.Echo "Description:" & conf.Description
  WScript.Echo "Name: " & conf.Name
  WScript.Echo "Scratch Directory:" & conf.ScratchDirectory
  WScript.Echo "Setting ID:" & conf.SettingID '获取不到
  WScript.Echo "Temp Directory: " & conf.TempDirectory
Next
