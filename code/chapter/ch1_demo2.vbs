' Option Explicit
' On Error Resume Next

' Dim objShell
' Dim regActivePCName,regPCName,regHostName
' Dim activePCName,PCName,HostName

regActivePCName = "HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName\ComputerName" 
   
regPCName = "HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName"

regHostName = "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\HostName" 

set objShell = CreateObject("WScript.Shell")   

activePCName = objShell.RegRead(regActivePCName)
PCName = objShell.RegRead(regPCName)
HostName = objShell.RegRead(regHostName)

' WScript.Echo "ActiveComputerName:" & activePCName 
' WScript.Echo "ComputerName:" & PCName 
' WScript.Echo "HostName:" & HostName

MsgBox("ActiveComputerName:" & activePCName)
MsgBox("ComputerName:" & PCName)
MsgBox("HostName:" & PCName)