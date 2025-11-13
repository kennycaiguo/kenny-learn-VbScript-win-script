'查看电脑桌面信息
Set wmi = GetObject("winmgmts:\\.\root\cimv2")

Set cols = wmi.ExecQuery("SELECT * FROM Win32_Desktop")

For Each dtop In cols
   WScript.Echo "Caption:" & dtop.Caption
   WScript.Echo "Description:" & dtop.Description
   WScript.Echo "Name:" & dtop.Name
   WScript.Echo "BorderWidth:" & dtop.BorderWidth
   WScript.Echo "Cool Switch:" & dtop.CoolSwitch
   WScript.Echo "GridGranularity:" & dtop.GridGranularity 
   WScript.Echo "IconTitleSize: " & dtop.IconTitleSize  
   WScript.Echo "IconTitleWrap: " & dtop.IconTitleWrap
   WScript.Echo "Pattern: " & dtop.Pattern
   WScript.Echo "SettingID: " & dtop.SettingID
   WScript.Echo "Wallpaper:"  & dtop.Wallpaper
   WScript.Echo "WallpaperStretched:" & dtop.WallpaperStretched 
   WScript.Echo "WallpaperTiled:" & dtop.WallpaperTiled  
   WScript.Echo "==========================================="
Next