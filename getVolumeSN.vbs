' getVolumeSN.vbs v1.3.0 by jimmy19990
' ==========================
' This script helps you to obtain Volume Serial Numbers of all your attached disk devices.
' URL: https://github.com/jimmy19990/USpider.vbs
'
strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")

For Each objItem In colItems
    WScript.Echo "Device ID: " & objItem.DeviceID & _
    vbcrlf & "Volume Name: " & objItem.VolumeName & _
    vbcrlf & "Volume Serial Number: " & objItem.VolumeSerialNumber
Next
