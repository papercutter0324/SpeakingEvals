Option Explicit

#If Mac Then
Public Function AreAppleScriptsInstalled(Optional ByVal resourcesFolder As String = vbNullString, Optional ByVal libraryScriptsFolder As String = vbNullString, Optional ByVal recheckStatus As Boolean = False) As Boolean
    Dim isAppleScriptInstalled As Boolean
    Dim isDialogToolkitInstalled As Boolean
    Dim statusHasBeenChecked As Boolean
    Dim scriptResult As Boolean
    
    If resourcesFolder = vbNullString Then
        resourcesFolder = GetDefaultFolderPaths("Resources")
    End If
    
    If libraryScriptsFolder = vbNullString Then
        libraryScriptsFolder = GetDefaultFolderPaths("Libraries")
    End If
    
    isAppleScriptInstalled = CheckForAppleScript()
    
    If isAppleScriptInstalled Then
        If Not recheckStatus Then CheckForAppleScriptUpdate

        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.DialogToolKitPlus.AttemptToLocate", libraryScriptsFolder)
        End If

        If Not recheckStatus Then
            ' When first opened, only check for Dialog Toolkit Plus if the folder has been previously created
            On Error Resume Next
            scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFolderExist", libraryScriptsFolder)
            On Error GoTo 0
            If scriptResult Then isDialogToolkitInstalled = CheckForDialogToolkit(resourcesFolder)
        Else
            isDialogToolkitInstalled = CheckForDialogToolkit(resourcesFolder)
        End If

        ' This may be a redundant message
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.DialogToolKitPlus.InstalledStatus", isDialogToolkitInstalled)
        End If

        If isDialogToolkitInstalled Then
            isDialogToolkitInstalled = CheckForDialogDisplayScript(resourcesFolder)
            If g_UserOptions.EnableLogging Then
                DebugAndLogging GetMsg("Debug.DialogToolKitPlus.AttemptToInstall", isDialogToolkitInstalled)
            End If
        End If
    Else
        isDialogToolkitInstalled = False
    End If

    SetVisibilityOfMacSettingsShapes isAppleScriptInstalled, isDialogToolkitInstalled

    AreAppleScriptsInstalled = isAppleScriptInstalled
End Function

Public Function AreEnhancedDialogsEnabled() As Boolean
    AreEnhancedDialogsEnabled = ThisWorkbook.Sheets("MacOS Users").Shapes("Button_EnhancedDialogs_Enable").Visible
End Function

Public Function CheckForAppleScript() As Boolean
    Dim appleScriptPath As String
    Dim appleScriptStatus As Boolean
    
    appleScriptPath = GetDefaultFolderPaths("Scripts") & APPLE_SCRIPT_FILE
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.AppleScript.AttemptToLocate", APPLE_SCRIPT_FILE, appleScriptPath)
    End If
    
    On Error Resume Next
    ' Instead of boolean, an InStr might be a better choice in case of having additional files in the target folder
    appleScriptStatus = (Dir$(appleScriptPath, vbDirectory) = APPLE_SCRIPT_FILE)
    On Error GoTo 0
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.AppleScript.InstalledStatus", appleScriptStatus)
    End If
    
    CheckForAppleScript = appleScriptStatus
End Function

Public Sub CheckForAppleScriptUpdate()
    Dim scriptFolder As String
    Dim destinationPath As String
    Dim currentScriptVersion As Long
    Dim downloadedScriptVersion As Long
    Dim appleScriptResult As Boolean
    Dim updateStepComplete As Boolean
    Dim renameParamStringStepOne As String
    Dim renameParamStringStepTwo As String
    Dim deleteParamString As String
    
    Const OLD_APPLE_SCRIPT As String = "SpeakingEvals-Old.scpt"
    Const TMP_APPLE_SCRIPT As String = "SpeakingEvals-Tmp.scpt"
    
    scriptFolder = GetDefaultFolderPaths("Scripts")
    destinationPath = scriptFolder & TMP_APPLE_SCRIPT
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.AppleScript.CheckForUpdate")
    End If
    
    updateStepComplete = DownloadFile("SpeakingEvals.scpt", "SpeakingEvals", destinationPath)
    If Not updateStepComplete Then
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.AppleScript.UnableToDownloadUpdate", APPLE_SCRIPT_FILE)
        End If
        GoTo ErrorHandler
    End If
    
    On Error GoTo ErrorHandler
    currentScriptVersion = AppleScriptTask(APPLE_SCRIPT_FILE, "GetScriptVersionNumber", vbNullString)
    downloadedScriptVersion = AppleScriptTask(TMP_APPLE_SCRIPT, "GetScriptVersionNumber", vbNullString)
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.AppleScript.VersionNumbers", currentScriptVersion, downloadedScriptVersion)
    End If
    
    If downloadedScriptVersion <= currentScriptVersion Then
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.AppleScript.LatestVersionInstalled")
        End If
        GoTo Cleanup
    End If
    
    renameParamStringStepOne = scriptFolder & APPLE_SCRIPT_FILE & APPLE_SCRIPT_SPLIT_KEY & scriptFolder & OLD_APPLE_SCRIPT
    updateStepComplete = AppleScriptTask(TMP_APPLE_SCRIPT, "RenameFile", renameParamStringStepOne)
    
    If updateStepComplete Then
        renameParamStringStepTwo = scriptFolder & TMP_APPLE_SCRIPT & APPLE_SCRIPT_SPLIT_KEY & scriptFolder & APPLE_SCRIPT_FILE
        updateStepComplete = AppleScriptTask(OLD_APPLE_SCRIPT, "RenameFile", renameParamStringStepTwo)
    End If
    
    If updateStepComplete Then
        deleteParamString = scriptFolder & OLD_APPLE_SCRIPT
        updateStepComplete = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", deleteParamString)
    End If
    On Error GoTo 0
    
    If Not updateStepComplete Then GoTo ErrorHandler
    
    If g_UserOptions.EnableLogging Then
        If updateStepComplete Then
            DebugAndLogging GetMsg("Debug.AppleScript.UpdateComplete")
        End If
    End If
    
Cleanup:
    If g_UserOptions.EnableLogging Then
        If updateStepComplete Then
            DebugAndLogging GetMsg("Debug.FileManagement.BeginCleanUp")
        End If
    End If
    
    On Error Resume Next
    updateStepComplete = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", scriptFolder & TMP_APPLE_SCRIPT)
    If updateStepComplete Then
        updateStepComplete = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", scriptFolder & TMP_APPLE_SCRIPT)
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.FileManagement.RemoveTemporaryFile", IIf(updateStepComplete, "Successful", "Failed"))
        End If
    End If
    
    updateStepComplete = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", scriptFolder & OLD_APPLE_SCRIPT)
    If updateStepComplete Then
        updateStepComplete = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", scriptFolder & OLD_APPLE_SCRIPT)
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.FileManagement.RemoveOldVersion", IIf(updateStepComplete, "Successful", "Failed"))
        End If
    End If
    On Error GoTo 0
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.FileManagement.FinishedCleanUp")
    End If
    Exit Sub
    
ErrorHandler:
    If g_UserOptions.EnableLogging Then
        If Err.Number <> 0 Then
            DebugAndLogging GetMsg("Debug.ErrorMessages.ErrorDuringUpdateProcess", Err.Number, Err.Description)
        End If
    End If
    GoTo Cleanup
End Sub

Public Function CheckForDialogToolkit(ByVal resourcesFolder As String) As Boolean
    Dim scriptResult As Boolean
    Dim libraryScriptsPath As String
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.DialogToolKitPlus.AttemptToLocate", resourcesFolder)
    End If
    
    On Error Resume Next
    libraryScriptsPath = AppleScriptTask(APPLE_SCRIPT_FILE, "CheckForScriptLibrariesFolder", "paramString")
    
    If libraryScriptsPath <> vbNullString Then
        scriptResult = RequestFileAndFolderAccess(resourcesFolder, libraryScriptsPath)
        
        If scriptResult Then
            scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "InstallDialogToolkitPlus", resourcesFolder)
        End If
    End If
    On Error GoTo 0
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.DialogToolKitPlus.InstalledStatus", scriptResult)
    End If
    
    CheckForDialogToolkit = scriptResult
End Function

Public Function CheckForDialogDisplayScript(ByVal resourcesFolder As String) As Boolean
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.DialogDisplayScript.AttemptToLocate")
    End If
        
    On Error Resume Next
    CheckForDialogDisplayScript = AppleScriptTask(APPLE_SCRIPT_FILE, "InstallDialogDisplayScript", resourcesFolder)
    On Error GoTo 0
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.CodeExecution.Status", vbTab, CheckForDialogDisplayScript)
    End If
End Function

Public Sub RemoveDialogToolKit(ByVal resourcesFolder As String)
    Dim scriptResult As Boolean
        
    If CheckForAppleScript() Then
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.DialogToolKitPlus.RemoveInstalledFile", resourcesFolder)
        End If
            
        On Error Resume Next
        scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "UninstallDialogToolkitPlus", resourcesFolder)
        On Error GoTo 0
            
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.CodeExecution.Result", INDENT_LEVEL_1, scriptResult)
        End If
    End If
End Sub

Public Sub RemindUserToInstallSpeakingEvalsScpt()
    DisplayMessage "Display.AppleScript.InstallReminder"
    MacOS_Users.Activate
End Sub

Public Function RequestFileAndFolderAccess(ByVal resourcesFolder As String, Optional ByVal filePath As Variant = vbNullString) As Boolean
    Dim workingFolder As Variant
    Dim excelTempFolder As Variant
    Dim powerpointTempFolder As Variant
    Dim filePermissionCandidates As Variant
    Dim pathToRequest As Variant
    Dim fileAccessGranted As Boolean
    Dim allAccessHasBeenGranted As Boolean
    Dim i As Long

    Select Case filePath
        Case vbNullString
            workingFolder = GetDefaultFolderPaths("Base")
            excelTempFolder = GetDefaultFolderPaths("Temp")
            powerpointTempFolder = Replace(excelTempFolder, "Excel", "PowerPoint")
            filePermissionCandidates = Array(workingFolder, resourcesFolder, excelTempFolder, powerpointTempFolder)
        Case Else
            filePermissionCandidates = Array(filePath)
    End Select

    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.FileManagement.FileAccessPermissionRequest.Message")
    End If

    For i = LBound(filePermissionCandidates) To UBound(filePermissionCandidates)
        pathToRequest = Array(filePermissionCandidates(i))
        fileAccessGranted = GrantAccessToMultipleFiles(pathToRequest)
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.FileManagement.FileAccessPermissionStatusDetailed.Message", filePermissionCandidates(i), IIf(fileAccessGranted, "granted", "denied"))
        End If
        allAccessHasBeenGranted = fileAccessGranted
        If Not fileAccessGranted Then Exit For
    Next i

    RequestFileAndFolderAccess = allAccessHasBeenGranted
End Function

Public Sub TriggerSystemAccessPrompt()
    Dim testPath As String
    Dim testFile As Integer
    Dim fileAccessGranted As Boolean
    
    testPath = GetDefaultFolderPaths("Base") & "ExcelPermissionTest.txt"
    
    On Error Resume Next
    testFile = FreeFile
    Open testPath For Output As #testFile
    Print #testFile, "Trigger MacOS prompt for Excel file and folder access."
    Close #testFile
    Kill testPath
    On Error GoTo 0
End Sub

Public Function SetFileAsExecutable(ByVal filePath As String) As Boolean
    Dim isExecutable As Boolean

    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.FileManagement.MarkAsExecutable", filePath)
    End If

    On Error Resume Next
    isExecutable = AppleScriptTask(APPLE_SCRIPT_FILE, "ChangeFilePermissions", "+x" & APPLE_SCRIPT_SPLIT_KEY & destinationPath)
    On Error GoTo 0

    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.FileManagement.MarkAsExecutableResult", IIf(isExecutable, "Successful", "Failed"))
    End If

    SetFileAsExecutable = isExecutable
End Function
#End If