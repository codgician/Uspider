strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")

For Each objItem In colItems

WScript.Echo "Device ID: " & objItem.DeviceID & _
vbcrlf & "Volume Serial Number: " & objItem.VolumeSerialNumber
Next
