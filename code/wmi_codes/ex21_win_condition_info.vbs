Set wmi = GetObject("winmgmts:\\.\root\cimv2")

Set cols = wmi.ExecQuery("SELECT * FROM Win32_Condition")

For Each cond In cols
   WScript.Echo "Caption:" & cond.Caption
   WScript.Echo "Description:" & cond.Description
   WScript.Echo "CheckID:" & cond.CheckID
   WScript.Echo "Condition:" & cond.Condition
   WScript.Echo "Feature:" & cond.Feature
   WScript.Echo "Level:" & cond.Level
   WScript.Echo "Name:" & cond.Name
   WScript.Echo "==========================================="
Next
