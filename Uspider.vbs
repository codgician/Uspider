' Uspider v2.0.0.alpha by codgician
' ====================================
' The brand new Uspider, more classy than ever.
' With great might comes great responsibility. DO NOT BE EVIL.
' Lincensed under MIT License.
' URL: https://github.com/codgician/USpider.vbs

Class Uspider
    Private objFileSystem, objWMIService, objWScriptShell, colEvents
    Private bool_logging, string_destFolder, bool_separateFolders, string_xcopyParameters, bool_isBlackList, string_customList
    
    ' Get parameters.
    Public Property Let logging(logOpt)
        bool_logging = logOpt
    End Property
 
    Public Property Let destFolder(destFolderOpt)
        string_destFolder = destFolderOpt
    End Property
    
    Public Property Let separateFolders(separateFoldersOpt)
        bool_separateFolders = separateFoldersOpt
    End Property
    
    Public Property Let xcopyParameters(xcopyParametersOpt)
        string_xcopyParameters = xcopyParametersOpt
    End Property
    
    Public Property Let isBlackList(isBlackListOpt)
        bool_isBlackList = isBlackListOpt
    End Property
    
    Public Property Let customList(customListOpt)
        string_customList = customListOpt
    End Property
    
    Private Sub Class_Initialize() 
        Set objFileSystem = CreateObject("Scripting.FileSystemObject")
        Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
        Set objWScriptShell = CreateObject("WScript.Shell")
        Set colEvents = objWMIService.ExecNotificationQuery _
            ("Select * From __InstanceOperationEvent Within 10 Where " _
                & "TargetInstance isa 'Win32_LogicalDisk'")
    End Sub

    Private Sub Class_Terminate() 
        Set objFileSystem = Nothing
        Set objWMIService = Nothing
        Set objWScriptShell = Nothing
        Set colEvents = Nothing
    End Sub
    
    Private Function checkList(objTargetDevice)
        ' Check if the device is in custom list.
        isIncluded = false
        If IsEmpty(bool_isBlackList) = false and IsEmpty(string_customList) = false Then
            If InStr(string_customList, objTargetDevice.VolumeSerialNumber) > 0 Then
                isIncluded = true
            End If
                        
            log "Custom List Found! isIncluded = " & isIncluded

             If isIncluded = bool_isBlackList Then
                log "Files on this device will be copied..."
                checkList = true
            Else
                log "Files on this device will not be copied..."
                checkList = false
            End If
        Else
            log "No list is detected. Skip checking..."
            checkList = true
        End If
    End Function
    
    Private Function xcopy(objTargetDevice, destDir)
        ' Initialize Destination Folder.
        If objFileSystem.folderExists(destDir) = false Then
            objFileSystem.CreateFolder destDir
        End If
        
        ' Initialize Work Folder.
        If bool_separateFolders = true Then
            workFolder = destDir + "\" + objTargetDevice.VolumeSerialNumber
        Else
            workFolder = destDir
        End If
                        
        If objFileSystem.folderExists(workFolder) = false Then
            objFileSystem.CreateFolder workFolder
        End If
        
        copyCommand = "cmd.exe /c xcopy " + objTargetDevice.DeviceId + "\* " + workFolder + " " + string_xcopyParameters
        objWScriptShell.Run(copyCommand), 0, false
        
        log "Copying Thread Started. Command: " & copyCommand
    End Function
    
    Private Function log(logText)
        ' Check if bool_logging is on.
        If bool_logging <> true Then
            Exit Function
        End If
        
        ' Initialize bool_logging.
        If objFileSystem.FileExists("Uspider_log.txt") = false Then
            Set objLogFile = objFileSystem.CreateTextFile("Uspider_log.txt")
            objLogFile.Close
        End If
        
        Set objLogFile = objFileSystem.OpenTextFile("Uspider_log.txt", 8, True)
        
        objLogFile.Write("[" & Now & "] " & logText) & vbcrlf
        
        objLogFile.Close
    End Function
    
    Private Function watchDog()
        log "Uspider is now started..."
        Do While True
            Set objEvent = colEvents.NextEvent
            Set objTargetDevice = objEvent.TargetInstance
            
            ' Check if the target device type is Removable Device (DriveType = 2).
            If objTargetDevice.DriveType = 2 Then
                Select Case objEvent.Path_.Class
                    ' Insert
                    Case "__InstanceCreationEvent"
                                
                        log "New Device Inserted. DeviceID: " & objTargetDevice.DeviceId & _
                            " | Label: " & objTargetDevice.VolumeName & " | SN:" & objTargetDevice.VolumeSerialNumber

                        If checkList(objTargetDevice) = true Then
                            xcopy objTargetDevice, string_destFolder
                        End If
                        
                    ' Remove
                    Case "__InstanceDeletionEvent"
                        log "Device Removed. DeviceID: " & objTargetDevice.DeviceId & _
                            " | Label: " & objTargetDevice.VolumeName & " | SN:" & objTargetDevice.VolumeSerialNumber
                End Select
            End If
        Loop
    End Function
    
    Public Function Init()
        log "Initialization Finished."
        watchDog()
    End Function
End Class

' --------------------
'  Main Script
' --------------------

Set Spider = New Uspider

' Configurations
Spider.logging = true
Spider.destFolder = "D:\Uspider"
Spider.separateFolders = true
Spider.xcopyParameters = "/d /e /r /y /h"
Spider.isBlackList = false
Spider.customList = ""

' Start spying!
Spider.Init()
