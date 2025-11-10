'查看电脑桌面显示器信息
Set wmi = GetObject("winmgmts:\\.\root\cimv2")

Set cols = wmi.ExecQuery("SELECT * FROM Win32_DesktopMonitor")

For Each dtop In cols
   WScript.Echo "Caption:" & dtop.Caption
   WScript.Echo "Description:" & dtop.Description
   WScript.Echo "DeviceID:" & dtop.DeviceID
   WScript.Echo "Availability :" & dtop.Availability 
   WScript.Echo "DisplayType:" & dtop.DisplayType 
   WScript.Echo "MonitorManufacturer:" & dtop.MonitorManufacturer
   WScript.Echo "MonitorType:" & dtop.MonitorType
   WScript.Echo "Name:" & dtop.Name
   WScript.Echo "PixelsPerXLogicalInch:" & dtop.PixelsPerXLogicalInch
   WScript.Echo "PixelsPerYLogicalInch:"  & dtop.PixelsPerYLogicalInch
   WScript.Echo "PowerManagementCapabilities:" & dtop.PowerManagementCapabilities
   WScript.Echo "PowerManagementSupported:" & dtop.PowerManagementSupported
   WScript.Echo "ScreenWidth:" & dtop.ScreenWidth
   WScript.Echo "ScreenHeight:" & dtop.ScreenHeight
   WScript.Echo "Status:" & dtop.Status
   WScript.Echo "==========================================="
Next