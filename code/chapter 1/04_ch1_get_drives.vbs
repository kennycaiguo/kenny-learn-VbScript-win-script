Option Explicit
On Error Resume Next
'DriveType = 3 fixed disk ,4->network 5=>CD Rom
Const DriveType =3 
Dim colDrives,drive 

set colDrives = GetObject("winmgmts:").ExecQuery _
    ("select DeviceID from Win32_LogicalDisk where DriveType=" & DriveType)

' travel
For Each drive In colDrives
    WScript.Echo drive.DeviceID
Next ' drive    