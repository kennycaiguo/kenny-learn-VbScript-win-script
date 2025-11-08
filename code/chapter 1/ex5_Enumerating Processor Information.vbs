' Enumerating Processor Information


On Error Resume Next
strInfo = ""
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")

For Each objItem in colItems
    strInfo = strInfo & "Address Width: " & objItem.AddressWidth & vbCrLf
    strInfo = strInfo & "Architecture: " & objItem.Architecture & vbCrLf
    strInfo = strInfo & "Availability: " & objItem.Availability & vbCrLf
    strInfo = strInfo & "CPU Status: " & objItem.CpuStatus & vbCrLf
    strInfo = strInfo & "Current Clock Speed: " & objItem.CurrentClockSpeed & vbCrLf
    strInfo = strInfo & "Data Width: " & objItem.DataWidth & vbCrLf
    strInfo = strInfo & "Description: " & objItem.Description & vbCrLf
    strInfo = strInfo & "Device ID: " & objItem.DeviceID & vbCrLf
    strInfo = strInfo & "External Clock: " & objItem.ExtClock & vbCrLf
    strInfo = strInfo & "Family: " & objItem.Family & vbCrLf
    strInfo = strInfo & "L2 Cache Size: " & objItem.L2CacheSize & vbCrLf
    strInfo = strInfo & "L2 Cache Speed: " & objItem.L2CacheSpeed & vbCrLf
    strInfo = strInfo & "Level: " & objItem.Level & vbCrLf
    strInfo = strInfo & "Load Percentage: " & objItem.LoadPercentage & vbCrLf
    strInfo = strInfo & "Manufacturer: " & objItem.Manufacturer & vbCrLf
    strInfo = strInfo & "Maximum Clock Speed: " & objItem.MaxClockSpeed & vbCrLf
    strInfo = strInfo & "Name: " & objItem.Name & vbCrLf
    strInfo = strInfo & "PNP Device ID: " & objItem.PNPDeviceID & vbCrLf
    strInfo = strInfo & "Processor ID: " & objItem.ProcessorId & vbCrLf
    strInfo = strInfo & "Processor Type: " & objItem.ProcessorType & vbCrLf
    strInfo = strInfo & "Revision: " & objItem.Revision & vbCrLf
    strInfo = strInfo & "Role: " & objItem.Role & vbCrLf
    strInfo = strInfo & "Socket Designation: " & objItem.SocketDesignation & vbCrLf
    strInfo = strInfo & "Status Information: " & objItem.StatusInfo & vbCrLf
    strInfo = strInfo & "Stepping: " & objItem.Stepping & vbCrLf
    strInfo = strInfo & "Unique Id: " & objItem.UniqueId & vbCrLf
    strInfo = strInfo & "Upgrade Method: " & objItem.UpgradeMethod & vbCrLf
    strInfo = strInfo & "Version: " & objItem.Version & vbCrLf
    strInfo = strInfo & "Voltage Caps: " & objItem.VoltageCaps & vbCrLf
Next

WScript.Echo strInfo