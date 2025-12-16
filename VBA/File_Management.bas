Option Explicit

#Const Windows = (Mac = 0)

Public Function ConvertToLocalPath(ByVal initialPath As Variant) As String
    Static convertedPath As String

    Dim i As Long
    
    ' Cloud storage apps like OneDrive sometimes complicate where/how files are saved. Below is a reference
    ' to track and help add support for additionalcloud storage providers.
    
    ' OneDrive
        ' Local Paths:      "/Users/" & Environ("USER") & "/Library/CloudStorage/OneDrive-Personal/"
        ' Returned Paths:   https://d.docs.live.net  AND  OneDrive://
        ' Procedure:        Trim everything before the 4th '/'
    ' iCloud
        ' Local Paths:      "/Users/" & Environ("USER") & "/Library/Mobile Documents/com~apple~CloudDocs/"
        ' Returned Paths:   N/A
        ' Procedure:        No trim required. ThisWorkbook.Path returns full local path
    ' Google Drive
        ' Local Paths:      "/Users/" & Environ("USER") & "/Library/CloudStorage/GoogleDrive-[user]@gmail.com/"
        ' Returned Paths:   N/A
        ' Procedure:        No trim required. ThisWorkbook.Path returns full local path
    
    If Left$(initialPath, 23) = "https://d.docs.live.net" Or Left$(initialPath, 11) = "OneDrive://" Then
        For i = 1 To 4
            initialPath = Mid$(initialPath, InStr(initialPath, "/") + 1)
        Next i
        
    #If Mac Then
        convertedPath = "/Users/" & Environ("USER") & "/Library/CloudStorage/OneDrive-Personal/" & initialPath
    #Else
        convertedPath = Environ$("OneDrive") & Application.PathSeparator & Replace(initialPath, "/", Application.PathSeparator)
    #End If
    Else
        convertedPath = initialPath
    End If
    
    ConvertToLocalPath = convertedPath
End Function

Public Function CheckForAndAttemptToCreateFolder(ByVal folderPath As String, Optional ByVal subFolderName As String = vbNullString, Optional ByVal clearContents As Boolean = False) As Boolean
    Dim requestPermissionsAgain As Boolean

    If Not DoesFolderExist(folderPath, True) Then
        CreateNewFolder folderPath
    #If Mac Then
        requestPermissionsAgain = True
    #End If
    ElseIf clearContents And subFolderName = vbNullString Then
        ClearFolder folderPath
    End If

    If subFolderName <> vbNullString Then
        EnsureTrailingPathSeparator folderPath
        folderPath = folderPath & subFolderName
        CheckForAndAttemptToCreateFolder folderPath, vbNullString, clearContents
    End If

    CheckForAndAttemptToCreateFolder = DoesFolderExist(folderPath, requestPermissionsAgain)
End Function

Public Sub CreateNewFolder(ByRef filePath As String)
    Dim finalCharIsSlash As Boolean
    
    If Right$(filePath, 1) = Application.PathSeparator Then
        filePath = Left$(filePath, Len(filePath) - 1)
        finalCharIsSlash = True
    End If

    On Error GoTo ErrorHandler
#If Mac Then
    Dim scriptResult As Boolean
    scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CreateFolder", filePath)
#Else
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder filePath
#End If
    On Error GoTo 0

    If finalCharIsSlash Then
        filePath = filePath & Application.PathSeparator
    End If
    
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        Case -2147024894 ' Folder already exists
            ' Optionally handle the case where the folder already exists
            Resume Next
        Case Else
            DebugAndLogging GetMsg("Debug.FileManagement.ErrorCreatingFolder", Err.Number, Err.Description)
    End Select
End Sub

Public Function DoesFolderExist(ByVal folderPath As String, Optional ByVal requestPermissions As Boolean = False) As Boolean
#If Mac Then
    Dim folderFound As Boolean
    
    On Error Resume Next
    folderFound = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFolderExist", folderPath)
    On Error GoTo 0
    
    If folderFound And requestPermissions Then
        folderFound = RequestFileAndFolderAccess(vbNullString, folderPath)
    End If

    DoesFolderExist = folderFound
#Else
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    DoesFolderExist = fso.folderExists(folderPath)
#End If
End Function

Private Sub ClearFolder(ByVal folderPath As String)
#If Mac Then
    Dim appleScriptResult As Boolean
    
    On Error Resume Next
    appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "ClearFolder", folderPath)
    On Error GoTo 0
#Else
    Dim fso As Object
    Dim fsoFolder As Object
    Dim fsoFile As Object
    Dim fileExt As String
    
    If Right$(folderPath, 1) = Application.PathSeparator Then
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsoFolder = fso.GetFolder(folderPath)
    
    For Each fsoFile In fsoFolder.Files
        fileExt = LCase$(fso.GetExtensionName(fsoFile.Name))
        If fileExt = "pptx" Or fileExt = "pdf" Or fileExt = "zip" Then
            fsoFile.Delete True
        End If
    Next fsoFile
#End If
End Sub

Public Function DoesFileExist(ByVal filePath As String, Optional ByVal requestPermissions As Boolean = False) As Boolean
#If Mac Then
    Dim fileExists As Boolean
    Dim filedeletion As Boolean
    Dim fileAvailabilityMsg As String
    
    On Error Resume Next
    fileExists = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", filePath)
    
    If fileExists Then
        fileAvailabilityMsg = AppleScriptTask(APPLE_SCRIPT_FILE, "VerifyFileIsAvailableLocally", filePath)
     
        If fileAvailabilityMsg <> "Ok" Then
            filedeletion = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", filePath)
            On Error GoTo 0
            
            fileExists = False
        End If
               
        If requestPermissions Then
            fileExists = RequestFileAndFolderAccess(vbNullString, filePath)
        End If
    End If
    
    DoesFileExist = fileExists
#Else
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    DoesFileExist = fso.fileExists(filePath)
#End If
End Function

Public Sub DeleteFile(ByVal filePath As String)
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.FileManagement.DeletingFile", filePath)
    End If
    
#If Mac Then
    Dim appleScriptResult As Boolean
    
    On Error Resume Next
    appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", filePath)
    
    If appleScriptResult Then
        appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", filePath)
    End If
    On Error GoTo 0
#Else
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    If fso.fileExists(filePath) Then
        fso.DeleteFile filePath, True
    End If
    On Error GoTo 0
#End If
End Sub

Public Function IsHashValid(ByVal filePath As String, ByVal fileType As String) As Boolean
    Dim expectedResult As String
    Dim hashResult As String
    
#If Windows Then
    Dim objShell As Object: Set objShell = CreateObject("WScript.Shell")
#End If

    If fileType <> "JSON" Then
        expectedResult = ReadValueFromDictionary(g_dictFileData, fileType, "hash")
    End If
    
    
    ' Code to handle intial JSON downloads in a new system or folder

    ' Ideally, query the online file. Add code to do this.
    
    ' If unable to query the online file or an error occurs,
    ' we just assume it is valid so that function can continue
    ' in an offline environment. Not ideal, but improvements
    ' and better considerations will be made later.
    
    Select Case True
        Case fileType = "JSON"
            IsHashValid = True
            Exit Function
        Case expectedResult = "Entry not found: " & fileType & ".hash"
            ' The errorlikely lies here
            IsHashValid = True
            Exit Function
        Case Else
            ' Query for the expected result, but this should really never
            ' trigger, since this file is downloaded first
    End Select
    
On Error GoTo ErrorHandler
#If Mac Then
    hashResult = AppleScriptTask(APPLE_SCRIPT_FILE, "GetMD5Hash", filePath)
    IsHashValid = (LCase$(hashResult) = LCase$(expectedResult))
#Else
    hashResult = objShell.Exec("cmd /c certutil -hashfile """ & filePath & """ MD5").StdOut.ReadAll
    If Left$(hashResult, 34) <> "CertUtil: -hashfile command FAILED" Then
        IsHashValid = (LCase$(expectedResult) = LCase$(Trim$(Split(hashResult, vbCrLf)(1))))
    End If
#End If
On Error GoTo 0
    
Cleanup:
#If Windows Then
    Set objShell = Nothing
#End If
    Exit Function
ErrorHandler:
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.FileManagement.HashCheckError", Err.Number, Err.Description)
    End If

    IsHashValid = False
    Resume Cleanup
End Function

Public Function IsFileLoadedFromTempDir() As Boolean
    Dim tempPath As String: tempPath = GetDefaultFolderPaths("Temp")
    Dim filePath As String: filePath = GetDefaultFolderPaths("Base")
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.Workbook.LoadedFromTempFolderCheck", filePath, tempPath)
    End If
    
    IsFileLoadedFromTempDir = (filePath = tempPath)
End Function

Public Function GetDefaultFolderPaths(ByVal requestedFolder As String) As String
    Static basePath As String
    
    Dim returnedPath As String
    
    If basePath = vbNullString Then
        basePath = ConvertToLocalPath(ThisWorkbook.Path)
        EnsureTrailingPathSeparator basePath
    End If
    
    Select Case requestedFolder
        Case "Base"
            returnedPath = basePath
        Case "Resources", "JSON"
            returnedPath = basePath & "Resources" & Application.PathSeparator
        Case "Logs"
            returnedPath = basePath & "Logs" & Application.PathSeparator
        Case "Temp"
            returnedPath = GetTempFilePath & Application.PathSeparator
#If Mac Then
        Case "Libraries"
            returnedPath = "/Users/" & Environ("USER") & "/Library/Script Libraries/"
        Case "Scripts"
            returnedPath = "/Users/" & Environ("USER") & "/Library/Application Scripts/com.microsoft.Excel/"
#Else
        Case "Windows Local Font Path"
            returnedPath = Environ$("LOCALAPPDATA") & "\Microsoft\Windows\Fonts\"
        Case "Windows System Font Path"
            returnedPath = Environ$("WINDIR") & "\Fonts\"
#End If
    End Select
    
    EnsureTrailingPathSeparator returnedPath
    
    GetDefaultFolderPaths = returnedPath
End Function

Public Sub EnsureTrailingPathSeparator(ByRef folderPath As String)
    If Right$(folderPath, 1) <> Application.PathSeparator Then
        folderPath = folderPath & Application.PathSeparator
    End If
End Sub

Public Function GetTempFilePath() As String
#If Mac Then
    Const ENV_TEMP_LABEL As String = "TMPDIR"
#Else
    Const ENV_TEMP_LABEL As String = "TEMP"
#End If
    
    Static tmpPath As String

    If tmpPath = vbNullString Then
        tmpPath = Environ$(ENV_TEMP_LABEL)
    End If
    
    GetTempFilePath = tmpPath
End Function

Public Function PrepareRequiredFile(ByVal fileName As String, ByVal filePath As String, ByVal fileType As String) As String
    Dim tmpFilePath As String
    Dim determinedPath As String

    tmpFilePath = GetDefaultFolderPaths("Temp") & fileName
    DeleteFile tmpFilePath

    If Not IsValidFilePresent(filePath, fileType) Then
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.FileManagement.DownloadNewCopy", fileName)
        End If

        If Not DownloadFile(fileName, fileType, filePath) Then
            If g_UserOptions.EnableLogging Then
                DebugAndLogging GetMsg("Debug.FileManagement.DownloadFailedCritical")
            End If
            
            PrepareRequiredFile = vbNullString
            Exit Function
        End If

        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.FileManagement.DownloadSuccessful", INDENT_LEVEL_2)
        End If
    End If

    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.FileManagement.ValidFileFound")
    End If
    
    If MoveFile(filePath, tmpFilePath) Then
        determinedPath = tmpFilePath
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.FileManagement.LoadTemporaryCopy")
        End If
    Else
        determinedPath = filePath
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.FileManagement.ErrorCreatingTemporaryCopy")
        End If
    End If

    PrepareRequiredFile = determinedPath
End Function

Private Function IsValidFilePresent(ByVal filePath As String, ByVal fileType As String) As Boolean
    Dim validFileFound As Boolean
    
    If DoesFileExist(filePath) Then
        validFileFound = IsHashValid(filePath, fileType)
    End If
    
    IsValidFilePresent = validFileFound
End Function

Public Function MoveFile(ByVal initialPath As String, ByVal destinationPath As String) As Boolean
    On Error GoTo MoveFailed
#If Mac Then
    MoveFile = AppleScriptTask(APPLE_SCRIPT_FILE, "CopyFile", initialPath & APPLE_SCRIPT_SPLIT_KEY & destinationPath)
    Exit Function
#Else
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.CopyFile initialPath, destinationPath, True
    MoveFile = True
    Exit Function
#End If
MoveFailed:
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.FileManagement.FileMoveFailed")
    End If
    
    MoveFile = False
End Function

Public Function SetSavePath(ByRef ws As Worksheet, Optional ByVal generateProcess As String = vbNullString, Optional ByVal resetRunStatus As Boolean = False) As String
    Dim workingPath As String
    Dim subFolderName As String
    Dim clearContents As Boolean
    
    Static subsequentRun As Boolean
    
    If Not subsequentRun Then
        subsequentRun = True
        clearContents = True
    End If
    
    If resetRunStatus Then
        subsequentRun = False
    End If
    
    workingPath = GetDefaultFolderPaths("Base") & GenerateSaveFolderName(ws) & Application.PathSeparator
    subFolderName = IIf(generateProcess = "Certificates", "Certificates", vbNullString)
    
    If Not CheckForAndAttemptToCreateFolder(workingPath, subFolderName, clearContents) Then
        SetSavePath = vbNullString
        Exit Function
    End If
    
    If subFolderName <> vbNullString Then
        workingPath = workingPath & subFolderName & Application.PathSeparator
    End If
    
    SetSavePath = workingPath
End Function

Private Function GenerateSaveFolderName(ByRef ws As Worksheet) As String
    Dim classIdentifier As String
    
    Select Case ws.Cells.Item(4, 3).Value
        Case "MonWed": classIdentifier = "MW - " & ws.Cells.Item(5, 3).Value
        Case "MonFri": classIdentifier = "MF - " & ws.Cells.Item(5, 3).Value
        Case "WedFri": classIdentifier = "WF - " & ws.Cells.Item(5, 3).Value
        Case "MWF": classIdentifier = "MWF - " & ws.Cells.Item(5, 3).Value
        Case "TTh": classIdentifier = "TTh - " & ws.Cells.Item(5, 3).Value
        Case "MWF (Class 1)": classIdentifier = "MWF-1"
        Case "MWF (Class 2)": classIdentifier = "MWF-2"
        Case "TTh (Class 1)": classIdentifier = "TTh-1"
        Case "TTh (Class 2)": classIdentifier = "TTh-2"
    End Select
    
    GenerateSaveFolderName = ws.Cells.Item(3, 3).Value & " (" & classIdentifier & ")"
End Function

Public Function SanitizeFileName(ByVal englishName As String) As String
    Dim invalidCharacters As Variant
    Dim ch As Variant
    
    invalidCharacters = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    
    For Each ch In invalidCharacters
        englishName = Replace(englishName, ch, "_")
    Next ch
    
    englishName = Trim$(englishName)
    
    Do While Right$(englishName, 1) = "."
        englishName = Left$(englishName, Len(englishName) - 1)
    Loop
    
    If Len(englishName) > 10 Then englishName = Trim$(Left$(englishName, 10))
    
    Do While Right$(englishName, 1) = "_"
        englishName = Left$(englishName, Len(englishName) - 1)
    Loop
    
    SanitizeFileName = Trim$(englishName)
End Function