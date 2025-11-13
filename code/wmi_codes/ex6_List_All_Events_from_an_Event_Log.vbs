' List All Events from an Event Log
On Error Resume Next 
strComputer = "."
strInfo=""
Set fs = CreateObject("Scripting.FileSystemObject")
Set ts = fs.OpenTextFile("./logdata.txt",2,True)
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colLoggedEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where Logfile = 'Application'")

For Each objEvent in colLoggedEvents
    ts.WriteLine  "Category: " & objEvent.Category & vbCrLf
    ts.WriteLine  "Computer Name: " & objEvent.ComputerName & vbCrLf
    ts.WriteLine  "Event Code: " & objEvent.EventCode & vbCrLf
    ts.WriteLine  "Message: " & objEvent.Message & vbCrLf
    ts.WriteLine  "Record Number: " & objEvent.RecordNumber & vbCrLf
    ts.WriteLine  "Source Name: " & objEvent.SourceName & vbCrLf
    ts.WriteLine  "Time Written: " & objEvent.TimeWritten & vbCrLf
    ts.WriteLine  "Event Type: " & objEvent.Type & vbCrLf
    ts.WriteLine  "User: " & objEvent.User & vbCrLf
    ts.WriteLine  "======================================" & vbCrLf
    
Next
set ts = Nothing
set fs = Nothing
