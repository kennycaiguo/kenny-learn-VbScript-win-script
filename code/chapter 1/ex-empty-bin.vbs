Const RECYCLE_BIN = &Ha&  
  
 Set FSO = CreateObject("Scripting.FileSystemObject")  
 Set ObjShell = CreateObject("Shell.Application")  
 Set ObjFolder = ObjShell.Namespace(RECYCLE_BIN)  
 Set ObjFolderItem = ObjFolder.Self  
  
  
Set colItems = ObjFolder.Items  
For Each objItem in colItems  
 If (objItem.Type = "File folder") Then  
 Else  
 FSO.DeleteFile(objItem.Path)  '先删除文件
 End If  
Next  
  
Set colItems = ObjFolder.Items  
For Each objItem in colItems  
 If (objItem.Type = "File folder") Then  
 Else  '这个ELse不能少,否则没有效果
 FSO.DeleteFolder(objItem.Path)  '再删除文件夹
 End If  
Next  

'这个程序的代码有点古怪,但是可以清空回收站