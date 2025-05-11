Option Explicit

#Const PRINT_DEBUG_MESSAGES = True
#If Mac Then
    Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
    Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
#End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File and Folder Management
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsFileLoadedFromTempDir() As Boolean
    Dim tempPath As String
    Dim filePath As String
    
    filePath = ThisWorkbook.FullName
    ConvertOneDriveToLocalPath filePath
    tempPath = GetTempFilePath(vbNullString)
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Checking if loaded from a temp folder." & vbNewLine & _
                    "    Current Path: " & filePath & vbNewLine & _
                    "    Temp Folder: " & tempPath
    #End If
    
    IsFileLoadedFromTempDir = (Left$(filePath, Len(tempPath)) = tempPath)
End Function

Public Sub ConvertOneDriveToLocalPath(ByRef selectedPath As Variant)
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
    
    If Left$(selectedPath, 23) = "https://d.docs.live.net" Or Left$(selectedPath, 11) = "OneDrive://" Then
        For i = 1 To 4
            selectedPath = Mid$(selectedPath, InStr(selectedPath, "/") + 1)
        Next
        
        #If Mac Then
            selectedPath = "/Users/" & Environ("USER") & "/Library/CloudStorage/OneDrive-Personal/" & selectedPath
        #Else
            selectedPath = Environ$("OneDrive") & "\" & Replace(selectedPath, "/", "\")
        #End If
    End If
End Sub

Public Sub CreateNewFolder(ByRef filePath As String)
    #If Mac Then
        Dim scriptResult As Boolean
    #Else
        Dim fso As Object
    #End If
    
    If Right$(filePath, 1) = Application.PathSeparator Then
        filePath = Left$(filePath, Len(filePath) - 1)
    End If

    #If Mac Then
        scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CreateFolder", filePath)
    #Else
        On Error Resume Next
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder filePath
        Set fso = Nothing
        On Error GoTo 0
    #End If

    If Right$(filePath, 1) <> Application.PathSeparator Then
        filePath = filePath & Application.PathSeparator
    End If
End Sub

Public Sub DeleteFile(ByVal filePath As String)
    #If Mac Then
        Dim appleScriptResult As Boolean
    #Else
        Dim fso As Object
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Deleting: " & filePath
    #End If
    
    #If Mac Then
        appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", filePath)
        If appleScriptResult Then appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", filePath)
    #Else
        On Error Resume Next
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        If fso.FileExists(filePath) Then fso.DeleteFile filePath, True
        filePath = Replace(filePath, " ", "%20")
        If fso.FileExists(filePath) Then fso.DeleteFile filePath, True
        
        Set fso = Nothing
        On Error GoTo 0
    #End If
End Sub

Public Sub DeleteExistingFolder(ByVal filePath As String)
    #If Mac Then
        Dim msgToDisplay As String
        Dim msgresult As Variant
        Dim scriptResult As Boolean
    #Else
        Dim fso As Object
    #End If
    
    
    #If Mac Then
        scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "ClearFolder", filePath)
    #Else
        Set fso = CreateObject("Scripting.FileSystemObject")

        If Right$(filePath, 1) = Application.PathSeparator Then
            filePath = Left$(filePath, Len(filePath) - 1)
        End If

        fso.DeleteFolder filePath, True
        Set fso = Nothing
    #End If
End Sub

Public Function DoesFolderExist(ByVal filePath As String) As Boolean
    #If Mac Then
        DoesFolderExist = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFolderExist", filePath)
    #Else
        DoesFolderExist = (Dir(filePath, vbDirectory) <> vbNullString)
    #End If
End Function

Public Function GenerateSaveFolderName(ByVal ws As Worksheet) As String
    Dim classIdentifier As String
    
    Select Case ws.Cells.Item(4, 3).Value
        Case "MonWed"
            classIdentifier = "MW - " & ws.Cells.Item(5, 3).Value
        Case "MonFri"
            classIdentifier = "MF - " & ws.Cells.Item(5, 3).Value
        Case "WedFri"
            classIdentifier = "WF - " & ws.Cells.Item(5, 3).Value
        Case "MWF"
            classIdentifier = "MWF - " & ws.Cells.Item(5, 3).Value
        Case "TTh"
            classIdentifier = "TTh - " & ws.Cells.Item(5, 3).Value
        Case "MWF (Class 1)": classIdentifier = "MWF-1"
        Case "MWF (Class 2)": classIdentifier = "MWF-2"
        Case "TTh (Class 1)": classIdentifier = "TTh-1"
        Case "TTh (Class 2)": classIdentifier = "TTh-2"
    End Select
    
    GenerateSaveFolderName = ws.Cells.Item(3, 3).Value & " (" & classIdentifier & ")"
End Function

Public Function GetTempFilePath(ByVal fileName As String) As String
    #If Mac Then
        GetTempFilePath = Environ("TMPDIR") & fileName
    #Else
        GetTempFilePath = Environ$("TEMP") & Application.PathSeparator & fileName
    #End If
End Function

Public Function MoveFile(ByVal initialPath As String, ByVal destinationPath As String) As Boolean
    Dim moveSuccessful As Boolean
    
    #If Mac Then
        ' No additional variables needed
    #Else
        Dim fso As Object
    #End If
    
    On Error Resume Next
    #If Mac Then
        moveSuccessful = AppleScriptTask(APPLE_SCRIPT_FILE, "CopyFile", initialPath & APPLE_SCRIPT_SPLIT_KEY & destinationPath)
    #Else
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFile initialPath, destinationPath
        moveSuccessful = (Err.Number = 0)
        Set fso = Nothing
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        If Not moveSuccessful Then
            Debug.Print "    Failed to move template to " & destinationPath
        End If
    #End If
    
    Err.Clear
    On Error GoTo 0
    MoveFile = moveSuccessful
End Function

Public Function SetSaveLocation(ByVal ws As Object, ByVal saveRoutine As String, ByVal resourcesFolder As String) As String
    Dim filePath As String
    
    #If Mac Then
        Dim permissionGranted As Boolean
        
        Const PERMISSION_GRANTED As String = "    Folder access granted. Continuing with process"
        Const PERMISSION_DENIED As String = "    Folder access denied. Cannot continue."
    #End If
    
    filePath = ThisWorkbook.Path & Application.PathSeparator & GenerateSaveFolderName(ws) & Application.PathSeparator
    ConvertOneDriveToLocalPath filePath
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Setting Save Path for Reports" & vbNewLine & _
                    "    Path: " & filePath
    #End If

    If DoesFolderExist(filePath) Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Path already exists. Clearing out old files."
        #End If
        DeleteExistingFolder filePath
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Creating save path"
    #End If
    
    CreateNewFolder filePath
    #If Mac Then
        permissionGranted = RequestFileAndFolderAccess(resourcesFolder, filePath)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print IIf(permissionGranted, PERMISSION_GRANTED, PERMISSION_DENIED)
        #End If
        If Not permissionGranted Then
            ' Add a savePath permission denied value
            SetSaveLocation = ""
            Exit Function
        End If
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Saving reports in: " & vbNewLine & _
                    "    " & filePath
    #End If
    
    SetSaveLocation = filePath
End Function

Public Function FindPathToArchiveTool(ByVal resourcesFolder As String, Optional ByRef archiverName As String = vbNullString) As String
    Dim downloadResult As Boolean
    Dim i As Long
    
    ' Declare OS specific variables, constants, and arrays
    #If Mac Then
        Dim scriptResultBoolean As Boolean
        
        Const RESOURCES_7ZIP_FILENAME As String = "7zz"
        Const RESOURCES_7ZIP_ARCHIVER_NAME As String = "Local 7zip"
    #Else
        Dim wshShell As Object
        Dim defaultPaths As Variant
        Dim archiverNames As Variant
        Dim exeNames As Variant
        Dim regKeys As Variant
        Dim regValue As String
    
        Const REG_KEY_7ZIP As String = "HKEY_LOCAL_MACHINE\SOFTWARE\7-Zip\Path"
        Const REG_KEY_7ZIP_32BIT As String = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\7-Zip\Path"
        
        Const ARCHIVER_NAME_7ZIP As String = "7Zip"
        
        Const EXE_NAME_7ZIP As String = "7z.exe"
        
        Const DEFAULT_PATH_7ZIP As String = "C:\Program Files\7-Zip\"
        Const DEFAULT_PATH_7ZIP_32Bit As String = "C:\Program Files (x86)\7-Zip\"
        
        Const RESOURCES_7ZIP_FILENAME As String = "7za.exe"
        Const RESOURCES_7ZIP_ARCHIVER_NAME As String = "Local 7zip"
    #End If
    
    ' Find available archive utility
    #If Mac Then
        scriptResultBoolean = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", resourcesFolder & Application.PathSeparator & RESOURCES_7ZIP_FILENAME)
        If scriptResultBoolean Then
            scriptResultBoolean = AppleScriptTask(APPLE_SCRIPT_FILE, "ChangeFilePermissions", "+x" & APPLE_SCRIPT_SPLIT_KEY & resourcesFolder & Application.PathSeparator & RESOURCES_7ZIP_FILENAME)
            If scriptResultBoolean Then FindPathToArchiveTool = resourcesFolder & Application.PathSeparator & RESOURCES_7ZIP_FILENAME
            Exit Function
        End If
    #Else
        defaultPaths = Array(DEFAULT_PATH_7ZIP, DEFAULT_PATH_7ZIP_32Bit)
        archiverNames = Array(ARCHIVER_NAME_7ZIP, ARCHIVER_NAME_7ZIP)
        exeNames = Array(EXE_NAME_7ZIP, EXE_NAME_7ZIP)
        regKeys = Array(REG_KEY_7ZIP, REG_KEY_7ZIP_32BIT)
        
        Set wshShell = CreateObject("WScript.Shell")
        
        ' First check default installation locations
        For i = LBound(defaultPaths) To UBound(defaultPaths)
            If Dir(defaultPaths(i) & exeNames(i)) <> vbNullString Then
                archiverName = archiverNames(i)
                FindPathToArchiveTool = defaultPaths(i) & exeNames(i)
                Exit Function
            End If
        Next i
        
        ' If not found, check the registry for paths to custom locations
        On Error Resume Next
        For i = LBound(regKeys) To UBound(regKeys)
            regValue = wshShell.RegRead(regKeys(i))
            If Err.Number = 0 And regValue <> vbNullString Then
                If Right$(regValue, 1) <> "\" Then regValue = regValue & "\"
                
                ' Verify executable exists before returning path
                If Dir(regValue & exeNames(i)) <> vbNullString Then
                    archiverName = archiverNames(i)
                    FindPathToArchiveTool = regValue & exeNames(i)
                    Exit Function
                End If
            End If
            Err.Clear
        Next i
        On Error GoTo 0
    #End If
    
    Download7Zip resourcesFolder, downloadResult
    
    If downloadResult Then
        archiverName = RESOURCES_7ZIP_ARCHIVER_NAME
        FindPathToArchiveTool = resourcesFolder & Application.PathSeparator & RESOURCES_7ZIP_FILENAME
    Else
        FindPathToArchiveTool = vbNullString
    End If
End Function

Public Function LocateTemplate(ByVal resourcesFolder As String, ByVal REPORT_TEMPLATE As String) As String
    Dim templatePath As String
    Dim tempTemplatePath As String
    Dim msgToDisplay As String
    Dim msgresult As Variant
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Loading Report Template"
    #End If
    
    templatePath = resourcesFolder & Application.PathSeparator & REPORT_TEMPLATE
    tempTemplatePath = GetTempFilePath(REPORT_TEMPLATE)
    
    DeleteFile tempTemplatePath ' Removing existing file to avoid problems overwriting

    If Not VerifyTemplateHash(templatePath) Then
        If Not DownloadReportTemplate(templatePath, resourcesFolder) Then
            msgToDisplay = "No template was found. Process canceled."
            msgresult = DisplayMessage(msgToDisplay, vbOKOnly, "Template Not Found", 150)
            LocateTemplate = vbNullString
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "    Unable to locate a copy of the template."
            #End If
            Exit Function
        End If
    End If

    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Valid template found"
    #End If
    
    If MoveFile(templatePath, tempTemplatePath) Then
        LocateTemplate = tempTemplatePath
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Loading temporary copy"
        #End If
    Else
        LocateTemplate = templatePath
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Failed to make a temporary copy. Using resources copy directly."
        #End If
    End If
End Function


Private Function VerifyTemplateHash(ByVal templatePath As String) As Boolean
    Const TEMPLATE_HASH As String = "AC1794BC6B04C8F18952D5A21A0BCEA4"
    
    #If Mac Then
        ' No extra variables required.
    #Else
        Dim objShell As Object
        Dim shellOutput As String
    #End If
    
    #If Mac Then
        VerifyTemplateHash = AppleScriptTask(APPLE_SCRIPT_FILE, "CompareMD5Hashes", templatePath & APPLE_SCRIPT_SPLIT_KEY & TEMPLATE_HASH)
    #Else
        If Dir(templatePath) <> vbNullString Then
            On Error GoTo ErrorHandler
            Set objShell = CreateObject("WScript.Shell")
            shellOutput = objShell.Exec("cmd /c certutil -hashfile """ & templatePath & """ MD5").StdOut.ReadAll
            VerifyTemplateHash = (LCase$(TEMPLATE_HASH) = LCase$(Trim$(Split(shellOutput, vbCrLf)(1))))
        Else
            VerifyTemplateHash = False
        End If
    #End If
CleanUp:
    #If Mac Then
    #Else
        Set objShell = Nothing
    #End If
    Exit Function
ErrorHandler:
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Error: " & Err.Number & " - " & Err.Description
    #End If
    VerifyTemplateHash = False
    Resume CleanUp
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

Public Sub DeletePDFs(ByVal targetFolder As String)
    #If Mac Then
    #Else
        Dim fso As Object
        Dim objFile As Object
        Dim objFolder As Object
    
        If Right$(targetFolder, 1) <> Application.PathSeparator Then targetFolder = targetFolder & Application.PathSeparator
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set objFolder = fso.GetFolder(targetFolder)
        
        On Error Resume Next
        For Each objFile In objFolder.Files
            If LCase$(fso.GetExtensionName(objFile.Name)) = "pdf" Then
                objFile.Delete True
                #If PRINT_DEBUG_MESSAGES Then
                    If Err.Number <> 0 Then
                        Debug.Print "Error deleting: " & objFile.Name & vbNewLine & _
                                    "Error: " & Err.Description
                        Err.Clear
                    End If
                #End If
            End If
        Next objFile
        On Error GoTo 0
    #End If
End Sub

Private Sub Download7Zip(ByVal resourcesFolder As String, ByRef downloadResult As Boolean)
    Dim destinationPath As String
    Dim downloadURL As String
    
    Const GIT_REPO_URL As String = "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/"
    
    #If Mac Then
        Dim scriptResultBoolean As Boolean
        
        Const FILE_NAME As String = "7zz"
    #Else
        Dim objWMI As Object
        Dim colProcessors As Object
        Dim objProcessor As Object
        Dim fileToDownload As String
        
        Const FILE_NAME As String = "7za.exe"
    #End If
    
    #If Mac Then
        destinationPath = resourcesFolder & Application.PathSeparator & FILE_NAME
        downloadURL = GIT_REPO_URL & FILE_NAME
        
        scriptResultBoolean = AppleScriptTask(APPLE_SCRIPT_FILE, "DownloadFile", destinationPath & APPLE_SCRIPT_SPLIT_KEY & downloadURL)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print IIf(scriptResultBoolean, "    Download successful.", "    Error: " & Err.Description)
        #End If
        
        If scriptResultBoolean Then
            downloadResult = RequestFileAndFolderAccess(resourcesFolder, destinationPath)
            scriptResultBoolean = AppleScriptTask(APPLE_SCRIPT_FILE, "ChangeFilePermissions", "+x" & APPLE_SCRIPT_SPLIT_KEY & destinationPath)
        End If
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    File access " & IIf(downloadResult, "granted.", "denied.")
        #End If
    #Else
        Set objWMI = GetObject("winmgmts:\\.\root\CIMV2")
        Set colProcessors = objWMI.ExecQuery("SELECT Architecture FROM Win32_Processor")
        
        For Each objProcessor In colProcessors
            Select Case objProcessor.architecture
                Case 0: fileToDownload = "7za(x86).exe"
                Case 9: fileToDownload = "7za(x64).exe"
                Case 12: fileToDownload = "7za(ARM).exe"
            End Select
        Next
        
        destinationPath = resourcesFolder & Application.PathSeparator & FILE_NAME
        downloadURL = GIT_REPO_URL & fileToDownload
        
        If Dir(destinationPath) <> vbNullString Then
            ' Add a hash check to verify the file
            downloadResult = True
            Exit Sub
        End If
        
        Select Case True
            Case CheckForCurl()
                downloadResult = DownloadUsingCurl(destinationPath, downloadURL)
            Case CheckForDotNet()
                downloadResult = DownloadUsingDotNet(destinationPath, downloadURL)
            Case Else
                downloadResult = False
        End Select
    #End If
End Sub

Private Function DownloadReportTemplate(ByVal templatePath As String, ByVal resourcesFolder As String) As Boolean
    Dim downloadResult As Boolean
    
    Const REPORT_TEMPLATE_URL As String = "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/SpeakingEvaluationTemplate.pptx"
    
    #If Mac Then
        On Error Resume Next
        downloadResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DownloadFile", templatePath & APPLE_SCRIPT_SPLIT_KEY & REPORT_TEMPLATE_URL)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print IIf(Err.Number = 0, "    Download successful.", "    Error: " & Err.Description)
        #End If
        
        If downloadResult Then downloadResult = RequestFileAndFolderAccess(resourcesFolder, templatePath)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    File access " & IIf(downloadResult, "granted.", "denied.")
        #End If
        On Error GoTo 0
    #Else
        If CheckForCurl() Then
            downloadResult = DownloadUsingCurl(templatePath, REPORT_TEMPLATE_URL)
        ElseIf CheckForDotNet() Then
            downloadResult = DownloadUsingDotNet(templatePath, REPORT_TEMPLATE_URL)
        Else
            downloadResult = False
        End If
    #End If
    
    If downloadResult Then
        DownloadReportTemplate = VerifyTemplateHash(templatePath)
    Else
        DownloadReportTemplate = False
    End If
End Function


#If Mac Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MacOS Only
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RequestFileAndFolderAccess(ByVal resourcesFolder As String, Optional ByVal filePath As Variant = "") As Boolean
    Dim workingFolder As Variant
    Dim excelTempFolder As Variant
    Dim powerpointTempFolder As Variant
    Dim filePermissionCandidates As Variant
    Dim pathToRequest As Variant
    Dim fileAccessGranted As Boolean
    Dim allAccessHasBeenGranted As Boolean
    Dim i As Long

    Select Case filePath
        Case ""
            workingFolder = ThisWorkbook.Path
            ConvertOneDriveToLocalPath workingFolder
            excelTempFolder = Environ("TMPDIR")
            powerpointTempFolder = Replace(excelTempFolder, "Excel", "PowerPoint")
            filePermissionCandidates = Array(workingFolder, resourcesFolder, excelTempFolder, powerpointTempFolder)
        Case Else
            ConvertOneDriveToLocalPath filePath ' Seems to be not needed?
            filePermissionCandidates = Array(filePath)
    End Select

    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Requesting access to: "
    #End If

    For i = LBound(filePermissionCandidates) To UBound(filePermissionCandidates)
        pathToRequest = Array(filePermissionCandidates(i))
        fileAccessGranted = GrantAccessToMultipleFiles(pathToRequest)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    " & filePermissionCandidates(i) & vbNewLine & _
                        "    Access granted: " & fileAccessGranted
        #End If
        allAccessHasBeenGranted = fileAccessGranted
        If Not fileAccessGranted Then Exit For
    Next i

    RequestFileAndFolderAccess = allAccessHasBeenGranted
End Function

#Else
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Windows Only
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DownloadUsingCurl(ByVal destinationPath As String, ByVal downloadURL As String) As Boolean
    Dim objShell As Object
    Dim fso As Object
    Dim downloadCommand As String
    
    On Error Resume Next
    Set objShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    downloadCommand = "cmd /c curl.exe -o """ & destinationPath & """ """ & downloadURL & """"
    objShell.Run downloadCommand, 0, True
    DownloadUsingCurl = fso.FileExists(destinationPath)
    
    #If PRINT_DEBUG_MESSAGES Then
        If Not DownloadUsingCurl Then Debug.Print "    curl download failed for " & downloadURL
    #End If
    On Error GoTo 0
End Function

Public Function DownloadUsingDotNet(ByVal destinationPath As String, ByVal downloadURL As String) As Boolean
    Dim xmlHTTP As Object
    Dim fileStream As Object
    
    On Error Resume Next
    Set xmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Set fileStream = CreateObject("ADODB.Stream")
    
    xmlHTTP.Open "Get", downloadURL, False
    xmlHTTP.Send
    
    If xmlHTTP.Status = 200 Then
        With fileStream
            .Open
            .Type = 1 ' Binary
            .Write xmlHTTP.responseBody
            .SaveToFile destinationPath, 2 ' Overwrite existing, if somehow present
            .Close
        End With
        DownloadUsingDotNet = True
    Else
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "HTTP request failed. Status: " & xmlHTTP.Status & " - " & xmlHTTP.StatusText
        #End If
        DownloadUsingDotNet = False
    End If
    On Error GoTo 0
End Function
#End If
