Option Explicit

#Const PRINT_DEBUG_MESSAGES = True
#If Mac Then
    Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
    Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
#End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File and Folder Management
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckForFolder(ByVal folderPath As String, Optional ByVal subFolderName As String = vbNullString, Optional ByVal clearContents As Boolean = False) As Boolean
    If Not DoesFolderExist(folderPath) Then
        CreateNewFolder folderPath
    ElseIf clearContents And subFolderName = vbNullString Then
        ClearFolder folderPath
    End If
    
    If subFolderName <> vbNullString Then
        If Right$(folderPath, 1) <> Application.PathSeparator Then
            folderPath = folderPath & Application.PathSeparator
        End If
        
        If Not DoesFolderExist(folderPath & subFolderName) Then
            CreateNewFolder (folderPath & subFolderName)
        ElseIf clearContents Then
            ClearFolder folderPath & subFolderName
        End If
        
        CheckForFolder = DoesFolderExist(folderPath & subFolderName)
    Else
        CheckForFolder = DoesFolderExist(folderPath)
    End If
End Function

Private Sub ClearFolder(ByVal folderPath As String)
    #If Mac Then
        Call AppleScriptTask(APPLE_SCRIPT_FILE, "ClearFolder", folderPath)
    #Else
        Dim fso As Object
        Dim fsoFolder As Object
        Dim fsoFile As Object
        Dim fileExt As String
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        'Step 1: Remove trailing path separator if present
        If Right$(folderPath, 1) = Application.PathSeparator Then
            folderPath = Left$(folderPath, Len(folderPath) - 1)
        End If
        
        ' Step 2: Iterate throught files and delete .pptx, .pdf, and .zip files found
        Set fsoFolder = fso.GetFolder(folderPath)
        
        For Each fsoFile In fsoFolder.Files
            fileExt = LCase$(fso.GetExtensionName(fsoFile.Name))
            If fileExt = "pptx" Or fileExt = "pdf" Or fileExt = "zip" Then
                fsoFile.Delete True
            End If
        Next fsoFile
        
        Set fsoFile = Nothing
        Set fsoFolder = Nothing
        Set fso = Nothing
    #End If
End Sub

Public Function ConvertOneDriveToLocalPath(ByVal initialPath As Variant) As String
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
            ConvertOneDriveToLocalPath = "/Users/" & Environ("USER") & "/Library/CloudStorage/OneDrive-Personal/" & initialPath
        #Else
            ConvertOneDriveToLocalPath = Environ$("OneDrive") & Application.PathSeparator & Replace(initialPath, "/", Application.PathSeparator)
        #End If
    Else
        ConvertOneDriveToLocalPath = initialPath
    End If
End Function

Public Sub CreateNewFolder(ByRef filePath As String)
    If Right$(filePath, 1) = Application.PathSeparator Then
        filePath = Left$(filePath, Len(filePath) - 1)
    End If

    #If Mac Then
        Dim scriptResult As Boolean
        
        On Error GoTo ErrorHandler
        scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CreateFolder", filePath)
        On Error GoTo 0
    #Else
        Dim fso As Object
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        On Error GoTo ErrorHandler
        fso.CreateFolder filePath
        On Error GoTo 0
        Set fso = Nothing
    #End If

    If Right$(filePath, 1) <> Application.PathSeparator Then
        filePath = filePath & Application.PathSeparator
    End If
    
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        Case -2147024894 ' Folder already exists
            ' Optionally handle the case where the folder already exists
            Resume Next
        Case Else
            Debug.Print "Error creating folder: " & Err.Description
    End Select
End Sub

Public Sub DeleteFile(ByVal filePath As String)
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print INDENT_LEVEL_1 & "Deleting: " & filePath
    #End If
    
    #If Mac Then
        Dim appleScriptResult As Boolean
        
        appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", filePath)
        
        If appleScriptResult Then
            appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", filePath)
        End If
    #Else
        Dim fso As Object
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        On Error Resume Next
        If fso.fileExists(filePath) Then
            fso.DeleteFile filePath, True
        End If
        
        filePath = Replace(filePath, " ", "%20")
        
        If fso.fileExists(filePath) Then
            fso.DeleteFile filePath, True
        End If
        On Error GoTo 0
        
        Set fso = Nothing
    #End If
End Sub

Public Function DoesFileExist(ByVal filePath As String) As Boolean
    #If Mac Then
        DoesFileExist = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", folderPath)
        If DoesFileExist Then
            DoesFileExist = RequestFileAndFolderAccess("", filePath)
        End If
    #Else
        DoesFileExist = (Dir(filePath) <> vbNullString)
    #End If
End Function

Public Function DoesFolderExist(ByVal folderPath As String) As Boolean
    #If Mac Then
        DoesFolderExist = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFolderExist", folderPath)
        If DoesFolderExist Then
            DoesFolderExist = RequestFileAndFolderAccess("", folderPath)
        End If
    #Else
        DoesFolderExist = (Dir(folderPath, vbDirectory) <> vbNullString)
    #End If
End Function

Public Function DownloadFileSuccessful(ByVal fileType As String, ByVal fileName As String, ByVal fileDestination As String) As Boolean
    Const GIT_REPO_URL As String = "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/"
    
    Dim repoSubfolder As String
    Dim downloadURL As String
    Dim downloadResult As Boolean
    
    Select Case UCase$(fileType)
        Case "TEMPLATE"
            repoSubfolder = "Templates/"
        Case "FONT"
            repoSubfolder = "Fonts/"
        Case "APPLESCRIPT"
            repoSubfolder = "AppleScript/"
        Case "7ZIP"
            repoSubfolder = "7zip/"
    End Select
    
    downloadURL = GIT_REPO_URL & repoSubfolder & fileName
    
    #If Mac Then
        downloadResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DownloadFile", fileDestination & APPLE_SCRIPT_SPLIT_KEY & downloadURL)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print IIf(Err.Number = 0, INDENT_LEVEL_1 & "Download successful.", INDENT_LEVEL_1 & "Error: " & Err.Description)
        #End If
        
        If downloadResult Then
            DownloadFileSuccessful = RequestFileAndFolderAccess("", fileDestination)
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print INDENT_LEVEL_1 & "File access " & IIf(downloadResult, "granted.", "denied.")
            #End If
        End If
    #Else
        Select Case True
            Case CheckForCurl()
                downloadResult = DownloadUsingCurl(fileDestination, downloadURL)
            Case CheckForDotNet()
                downloadResult = DownloadUsingDotNet(fileDestination, downloadURL)
            Case Else
                DownloadFileSuccessful = False
                Exit Function
        End Select
    #End If
    
    If downloadResult Then
        DownloadFileSuccessful = IsHashValid(fileDestination, fileName)
    Else
        DownloadFileSuccessful = False
    End If
End Function

Public Function FindPathToArchiveTool(ByVal resourcesFolder As String, Optional ByRef archiverName As String = vbNullString) As String
    Dim downloadResult As Boolean
    Dim i As Long
    
    #If Mac Then
        Dim scriptResultBoolean As Boolean
        
        Const RESOURCES_7ZIP_FILENAME As String = "7zz"
        Const RESOURCES_7ZIP_ARCHIVER_NAME As String = "Local 7zip"
        
        scriptResultBoolean = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", resourcesFolder & Application.PathSeparator & RESOURCES_7ZIP_FILENAME)
        
        If scriptResultBoolean Then
            scriptResultBoolean = AppleScriptTask(APPLE_SCRIPT_FILE, "ChangeFilePermissions", "+x" & APPLE_SCRIPT_SPLIT_KEY & resourcesFolder & Application.PathSeparator & RESOURCES_7ZIP_FILENAME)
            
            If scriptResultBoolean Then
                FindPathToArchiveTool = resourcesFolder & Application.PathSeparator & RESOURCES_7ZIP_FILENAME
            End If
            
            Exit Function
        End If
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

Private Function GenerateSaveFolderName(ByVal ws As Worksheet) As String
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

Private Function GetKnownGoodHash(ByVal fileName As String) As String
    Select Case fileName
        Case "7za(x86).exe"
            GetKnownGoodHash = "86D2E800B12CE5DA07F9BD2832870577"
        Case "7za(x64).exe"
            GetKnownGoodHash = "C58A4193BAC738B1A88ACAD9C6A57356"
        Case "7za(ARM).exe"
            GetKnownGoodHash = "3DCEBD415EC47C5EF080C13FAB5E15A2"
        Case "7zz"
            GetKnownGoodHash = "80B9D6E9761AECE7F8AC784491FC3B6A"
        Case "SpeakingEvaluationTemplate.pptx"
            GetKnownGoodHash = "AC1794BC6B04C8F18952D5A21A0BCEA4"
        Case "CertificateTemplate.pptx"
            GetKnownGoodHash = "BA41D2FDAE6F69A3FACBEC6BDF815D18"
        Case "just-another-hand.regular.ttf"
            GetKnownGoodHash = "2FBCF17635776DDB2692BF320838386C"
        Case "KakaoBigSans-Regular.ttf"
            GetKnownGoodHash = "01E7D95AE15377CB6747F824F1F6E9DB"
        Case "KakaoBigSans-Bold.ttf"
            GetKnownGoodHash = "FE621CE00147AADBD3E00134F38D0D86"
        Case "KakaoBigSans-ExtraBold.ttf"
            GetKnownGoodHash = "5427B26E380AC73E97A8E1B2CD1D108C"
        Case "DialogDisplay.scpt"
            GetKnownGoodHash = "33E05023335053FADC87B50900935E5E"
        Case "Dialog_Toolkit.zip"
            GetKnownGoodHash = "DB64101A9F28BA7C4D708FAAB760415C"
        Case "SpeakingEvals.scpt"
            GetKnownGoodHash = "68E2A7D937B9A145C15E823C45CE6E15"
        Case "mySignature.png"
            GetKnownGoodHash = "D3803794383425C34A11673C00033E85"
    End Select
End Function

Public Function GetTempFilePath(ByVal fileName As String) As String
    #If Mac Then
        GetTempFilePath = Environ("TMPDIR") & fileName
    #Else
        GetTempFilePath = Environ$("TEMP") & Application.PathSeparator & fileName
    #End If
End Function

Public Function IsFileLoadedFromTempDir() As Boolean
    Dim tempPath As String
    Dim filePath As String
    
    filePath = ConvertOneDriveToLocalPath(ThisWorkbook.fullName)
    tempPath = GetTempFilePath(vbNullString)
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Checking if loaded from a temp folder." & vbNewLine & _
                    INDENT_LEVEL_1 & "Current Path: " & filePath & vbNewLine & _
                    INDENT_LEVEL_1 & "Temp Folder: " & tempPath
    #End If
    
    IsFileLoadedFromTempDir = (Left$(filePath, Len(tempPath)) = tempPath)
End Function

Public Function IsHashValid(ByVal filePath As String, ByVal fileName As String) As Boolean
    #If Mac Then
        validateHash = AppleScriptTask(APPLE_SCRIPT_FILE, "CompareMD5Hashes", filePath & APPLE_SCRIPT_SPLIT_KEY & GetKnownGoodHash(fileName))
    #Else
        Dim objShell As Object
        Dim hashResult As String
        
        On Error GoTo ErrorHandler
        Set objShell = CreateObject("WScript.Shell")
        On Error GoTo 0
        
        hashResult = objShell.Exec("cmd /c certutil -hashfile """ & filePath & """ MD5").StdOut.ReadAll
        IsHashValid = (LCase$(GetKnownGoodHash(fileName)) = LCase$(Trim$(Split(hashResult, vbCrLf)(1))))
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
    IsHashValid = False
    Resume CleanUp
End Function

Private Function IsValidFilePresent(ByVal filePath As String, ByVal fileName As String) As Boolean
    IsValidFilePresent = DoesFileExist(filePath) And IsHashValid(filePath, fileName)
End Function

Public Function LocateRequiredFile(ByVal fileName As String, ByVal filePath As String, ByVal fileType As String) As String
    Dim tempFilePath As String
    
    ' Step 1: Set temporary file path
    tempFilePath = GetTempFilePath(fileName)
    
    ' Step 2: Remove existing temp file, if present, to avoid overwriting errors
    DeleteFile tempFilePath
    
    ' Step 3: Attempt to validate local copy
    If Not IsValidFilePresent(filePath, fileName) Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print INDENT_LEVEL_1 & "Valid " & fileName & " not found. Attempting to download."
        #End If
        
        ' Step 3a: Download new copy if missing or invalid
        If Not DownloadFileSuccessful(fileType, fileName, filePath) Then
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print INDENT_LEVEL_2 & "Download failed. Terminiating file creation."
            #End If
            
            LocateRequiredFile = vbNullString
            Exit Function
        End If
        
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print INDENT_LEVEL_2 & "Download successful."
        #End If
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print INDENT_LEVEL_1 & "Valid copy found."
    #End If
    
    ' Step 4: Make a temporary copy to work with
    If MoveFile(filePath, tempFilePath) Then
        LocateRequiredFile = tempFilePath
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print INDENT_LEVEL_1 & "Loading temporary copy"
        #End If
    Else
        ' Step 4a: Use local copy if unable to make a temporary copy
        LocateRequiredFile = filePath
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print INDENT_LEVEL_1 & "Failed to make a temporary copy. Using resources copy directly."
        #End If
    End If
End Function

Private Function MoveFile(ByVal initialPath As String, ByVal destinationPath As String) As Boolean
    Dim moveSuccessful As Boolean
    
    #If Mac Then
        moveSuccessful = AppleScriptTask(APPLE_SCRIPT_FILE, "CopyFile", initialPath & APPLE_SCRIPT_SPLIT_KEY & destinationPath)
    #Else
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        On Error Resume Next
        fso.CopyFile initialPath, destinationPath
        On Error GoTo 0
        
        moveSuccessful = (Err.Number = 0)
        Err.Clear
        Set fso = Nothing
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        If Not moveSuccessful Then
            Debug.Print INDENT_LEVEL_1 & "Failed to move template to " & destinationPath
        End If
    #End If
    
    MoveFile = moveSuccessful
End Function

Public Function SetSavePath(ByVal ws As Worksheet, Optional ByVal subFolderName As String = vbNullString) As String
    Dim workingPath As String
    
    ' Step 1: Set initial save location based on class days and time
    workingPath = ConvertOneDriveToLocalPath(ThisWorkbook.Path & Application.PathSeparator & GenerateSaveFolderName(ws) & Application.PathSeparator)
    
    If Not CheckForFolder(workingPath, subFolderName, True) Then
        SetSavePath = vbNullString
        Exit Function
    End If
    
    ' Step 2: Handle subfolder
    If subFolderName <> vbNullString Then
        workingPath = workingPath & subFolderName & Application.PathSeparator
    End If
    

    SetSavePath = workingPath
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
    
    Const GIT_REPO_URL As String = "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/7zip/"
    
    #If Mac Then
        Dim scriptResultBoolean As Boolean
        
        Const FILE_NAME As String = "7zz"
        
        destinationPath = resourcesFolder & Application.PathSeparator & FILE_NAME
        downloadURL = GIT_REPO_URL & FILE_NAME
        
        scriptResultBoolean = AppleScriptTask(APPLE_SCRIPT_FILE, "DownloadFile", destinationPath & APPLE_SCRIPT_SPLIT_KEY & downloadURL)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print IIf(scriptResultBoolean, INDENT_LEVEL_1 & "Download successful.", INDENT_LEVEL_1 & "Error: " & Err.Description)
        #End If
        
        If scriptResultBoolean Then
            downloadResult = RequestFileAndFolderAccess(resourcesFolder, destinationPath)
            scriptResultBoolean = AppleScriptTask(APPLE_SCRIPT_FILE, "ChangeFilePermissions", "+x" & APPLE_SCRIPT_SPLIT_KEY & destinationPath)
        End If
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print INDENT_LEVEL_1 & "File access " & IIf(downloadResult, "granted.", "denied.")
        #End If
    #Else
        Dim objWMI As Object
        Dim colProcessors As Object
        Dim objProcessor As Object
        Dim fileToDownload As String
        
        Const FILE_NAME As String = "7za.exe"
        
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
            workingFolder = ConvertOneDriveToLocalPath(ThisWorkbook.Path)
            excelTempFolder = Environ("TMPDIR")
            powerpointTempFolder = Replace(excelTempFolder, "Excel", "PowerPoint")
            filePermissionCandidates = Array(workingFolder, resourcesFolder, excelTempFolder, powerpointTempFolder)
        Case Else
            filePath = ConvertOneDriveToLocalPath(filePath) ' Seems to be not needed?
            filePermissionCandidates = Array(filePath)
    End Select

    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Requesting access to: "
    #End If

    For i = LBound(filePermissionCandidates) To UBound(filePermissionCandidates)
        pathToRequest = Array(filePermissionCandidates(i))
        fileAccessGranted = GrantAccessToMultipleFiles(pathToRequest)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print INDENT_LEVEL_1 & "" & filePermissionCandidates(i) & vbNewLine & _
                        INDENT_LEVEL_1 & "Access granted: " & fileAccessGranted
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
    DownloadUsingCurl = fso.fileExists(destinationPath)
    
    #If PRINT_DEBUG_MESSAGES Then
        If Not DownloadUsingCurl Then Debug.Print INDENT_LEVEL_1 & "curl download failed for " & downloadURL
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
