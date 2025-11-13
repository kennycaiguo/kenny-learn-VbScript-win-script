Set wmi = GetObject("winmgmts:\\.\root\cimv2")

Set cols = wmi.ExecQuery("SELECT * FROM Win32_CurrentTime")

For Each dtObj In cols
 WScript.Echo "Current Date:" & dtObj.Year & "-" & dtObj.Month & "-" & dtObj.Day
 WScript.Echo "Current Time:" & dtObj.Hour & ":" & dtObj.Minute & ":" & dtObj.second '小时数不正确
 WScript.Echo "Day Of Week:" & dtObj.DayOfWeek
 WScript.Echo "Quarter:" & dtObj.Quarter '不正确
 
 WScript.Echo "=============================="
Next