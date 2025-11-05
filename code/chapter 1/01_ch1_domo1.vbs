set objShell = CreateObject("Shell.Application")
set colTools = objShell.Namespace(47).Items

For Each objTool In colTools
    WScript.Echo objTool
Next ' objTool