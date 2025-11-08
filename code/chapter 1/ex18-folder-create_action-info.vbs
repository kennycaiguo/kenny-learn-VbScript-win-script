Set wmi = GetObject("winmgmts:\\.\root\cimv2")

Set cols = wmi.ExecQuery("SELECT * FROM Win32_CreateFolderAction")

For Each action In cols
 WScript.Echo "Name:" & action.Name '获取不到
 WScript.Echo "Action ID:" & action.ActionID
 WScript.Echo "Description:" & action.Description
 WScript.Echo "Directory:" & action.DirectoryName
 WScript.Echo "=============================="
Next