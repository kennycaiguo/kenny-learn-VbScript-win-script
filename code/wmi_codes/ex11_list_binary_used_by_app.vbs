On Error Resume Next
strcomputer = "."
Set fs = CreateObject("Scripting.FileSystemObject")
Set ts = fs.OpenTextFile("binary.txt",2,True)
Set wmiObj = GetObject _
  ("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strcomputer & "\root\cimv2")

Set colItems = wmiObj.ExecQuery("select * from Win32_Binary")

For Each item In colItems
    ts.WriteLine "Name:" & item.Name
    ts.WriteLine "Product Code:" & item.ProductCode
    ts.WriteLine
Next ' item

ts.Close()
set fs=Nothing