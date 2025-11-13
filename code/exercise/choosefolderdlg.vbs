Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.BrowseForFolder(0, "请选择一个文件夹", 0)
If Not objFolder Is Nothing Then
    MsgBox "您选择的文件夹是：" & objFolder.Self.Path
End If