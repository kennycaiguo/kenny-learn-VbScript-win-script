'查看USBController的下游集线器信息
Set wmi = GetObject("winmgmts:\\.\root\cimv2")

Set cols = wmi.ExecQuery("SELECT * FROM Win32_ControllerHasHub")

For Each conn In cols
   WScript.Echo "Antecedent:" & conn.Antecedent
   WScript.Echo "Dependent:" & conn.Dependent
   WScript.Echo "==========================================="
Next