Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.OpenTextFile("test.txt",2,True)
ts.Write "Hello,VBS"
ts.Close()
Set fso = Nothing
