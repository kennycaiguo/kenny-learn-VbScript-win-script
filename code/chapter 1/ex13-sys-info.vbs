'显示电脑系统信息
Set wmi = GetObject("winmgmts:\\.\root\cimv2")

Set cols = wmi.ExecQuery("SELECT * FROM Win32_ComputerSystem")

For Each sys In cols
 WScript.Echo "Caption:" & sys.Caption
 WScript.Echo "Domain:" & sys.Domain
 WScript.Echo "Description:" & sys.Description
 WScript.Echo "Time zone:" & sys.CurrentTimeZone
 WScript.Echo "DNS Host Name:" & sys.DNSHostName
 WScript.Echo "Domain Role:" & sys.DomainRole
 WScript.Echo "Name:" & sys.Name
 WScript.Echo "Model:" & sys.Model
 WScript.Echo "Status:" & sys.Status
 WScript.Echo "System Family:" & sys.SystemFamily
 WScript.Echo "User name:" & sys.UserName
 WScript.Echo "Work Group:" & sys.Workgroup
 WScript.Echo "Memory:" & sys.TotalPhysicalMemory
 WScript.Echo "System Type:" & sys.SystemType
 WScript.Echo "Startup Option:" & sys.SystemStartupOptions '获取不到
 WScript.Echo "System Startup Settings: " & sys.SystemStartupSetting '获取不到
 WScript.Echo "System Sku Number:" & sys.SystemSKUNumber 'ok
 WScript.Echo "System Startup Delay:" & sys.SystemStartupDelay '获取不到
 WScript.Echo "SupportContactDescription :" & sys.SupportContactDescription '获取不到
 Next
