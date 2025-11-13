On Error Resume Next
strcomputer = "."
Set fs = CreateObject("Scripting.FileSystemObject")
Set ts = fs.OpenTextFile("homenet.txt",2,True)
Set wmiObj = GetObject _
  ("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strcomputer & "\root\microsoft\homenet")

Set colItems = wmiObj.ExecQuery("select * from HNet_Connection")

For Each item In colItems
    ts.WriteLine "GUID: " & item.GUID
    ts.WriteLine "Is LAN Connection: " & item.IsLANConnection
    ts.WriteLine "Name: " & item.Name
    ts.WriteLine "Phone Book Path: " & item.PhoneBookPath
    ts.WriteLine
Next ' item

ts.Close()
set fs=Nothing