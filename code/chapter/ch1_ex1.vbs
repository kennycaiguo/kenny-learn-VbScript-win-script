Option Explicit
On Error Resume Next

Dim regAutoReboot,regMindumpDir,regLogEvent,regDumpFile
Dim autoReboot,minidumpDir,logEvent,dumpFile
Dim objShell

regAutoReboot = "HKLM\SYSTEM\CurrentControlSet\Control\CrashControl\AutoReboot"
regMindumpDir = "HKLM\SYSTEM\CurrentControlSet\Control\CrashControl\MinidumpDir"
regLogEvent = "HKLM\SYSTEM\CurrentControlSet\Control\CrashControl\LogEvent"
regDumpFile = "HKLM\SYSTEM\CurrentControlSet\Control\CrashControl\DumpFile"

set objShell = CreateObject("WScript.Shell")

autoReboot = objShell.RegRead(regAutoReboot)
minidumpDir = objShell.RegRead(regMindumpDir)
logEvent   = objShell.RegRead(regLogEvent)
dumpFile  = objShell.RegRead(regDumpFile)
WScript.Echo "AutoReboot:" & autoReboot
WScript.Echo "MinidumpDir: " & minidumpDir
WScript.Echo "LogEvent:" & logEvent
WScript.Echo "DumpFile: " & dumpFile