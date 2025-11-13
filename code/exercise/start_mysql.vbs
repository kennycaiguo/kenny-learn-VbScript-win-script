'下面的代码是可以启动mysql服务的但是需要以管理员权限打开cmd窗口
set wsobj = CreateObject("WScript.Shell")
wsobj.Run "sc config MySQL start=auto"
wsobj.Run "net start MySQL"
