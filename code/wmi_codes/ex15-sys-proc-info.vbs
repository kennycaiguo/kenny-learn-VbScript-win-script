Set wmi = GetObject("winmgmts:\\.\root\cimv2")

Set cols = wmi.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")

For Each proc In cols
  WScript.Echo "Name:" & proc.Name
  WScript.Echo "Caption:" & proc.Caption
  WScript.Echo "Description:" & proc.Description
  WScript.Echo "Identifying Number:" & proc.IdentifyingNumber
  WScript.Echo "Sku Number:" & proc.SKUNumber '获取不到
  WScript.Echo "UUID:" & proc.UUID
  WScript.Echo "Vendor:" & proc.Vendor
  WScript.Echo "Version:" & proc.Version '获取不到
Next
