Option Explicit

#Const Windows = (Mac = 0)

Public Function ProcessJsonUpdate(ByVal jsonFileName As String, ByVal dictPath As String, ByVal currentVersion As Long) As Boolean
    Dim tmpDictionary As New Dictionary
    Dim downloadPath  As String
    Dim latestVersion As Long
    Dim updateResult  As Boolean

    downloadPath = GetDefaultFolderPaths("Temp") & jsonFileName
    If DownloadFile(jsonFileName, "JSON", downloadPath) Then
        LoadValuesFromJson LoadDataFromJson(downloadPath), vbNullString, tmpDictionary
        latestVersion = CLng(tmpDictionary("Version.Number"))
        If currentVersion < latestVersion Then
            DeleteFile dictPath
            updateResult = MoveFile(downloadPath, dictPath)
        End If

        DeleteFile downloadPath
    End If
    
    ProcessJsonUpdate = updateResult
End Function

Public Function DownloadFile(ByVal fileName As String, ByVal fileType As String, ByVal fileDestination As String) As Boolean
    Dim downloadURL As String
    Dim downloadResult As Boolean
    
    downloadURL = GetDownloadUrl(fileType, fileName)

#If Mac Then
    On Error Resume Next
    downloadResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DownloadFile", fileDestination & APPLE_SCRIPT_SPLIT_KEY & downloadURL)
    On Error GoTo 0
    
    If g_UserOptions.EnableLogging Then
        If downloadResult Then
            DebugAndLogging GetMsg("Debug.FileManagement.DownloadSuccessful", vbTab)
        Else
            DebugAndLogging GetMsg("Debug.FileManagement.DownloadFailed", Err.Number, Err.Description)
        End If
    End If
    
    If downloadResult Then
        downloadResult = RequestFileAndFolderAccess(vbNullString, fileDestination)
        
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.FileManagement.FileAccessPermissionStatus", IIf(downloadResult, "granted.", "denied."))
        End If
    End If
        
    If downloadResult Then
        downloadResult = IsHashValid(fileDestination, fileType)
    End If
#Else
    Dim i As Long
    
    Do While i < 3
        i = i + 1
        
        If DownloadUsingHttpRequest(fileDestination, downloadURL) Then
            If IsHashValid(fileDestination, fileType) Then
                downloadResult = True
                Exit Do
            End If
        End If
    Loop
#End If
    
    If Not downloadResult Then
        DisplayMessage "Display.ErrorMessages.DownloadFile", fileName
    End If
    
    DownloadFile = downloadResult
End Function

Public Function DownloadUsingHttpRequest(ByVal destinationPath As String, ByVal downloadURL As String) As Boolean
#If Windows Then
    Dim xmlHTTP As Object: Set xmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim fileStream As Object: Set fileStream = CreateObject("ADODB.Stream")
    
    On Error GoTo ErrorHandler

    xmlHTTP.Open "GET", downloadURL, False
    xmlHTTP.Send
    
    If xmlHTTP.status = 200 Then
        With fileStream
            .Type = 1 'adTypeBinary
            .Open
            .write xmlHTTP.responseBody
            .SaveToFile destinationPath, 2 'adSaveCreateOverWrite
            .Close
        End With

        DownloadUsingHttpRequest = True
    Else
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.FileManagement.DownloadFailedHttp", xmlHTTP.status, xmlHTTP.StatusText)
        End If

        DownloadUsingHttpRequest = False
    End If

    Exit Function
ErrorHandler:
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.FileManagement.DownloadFailedHttp", Err.Number, Err.Description)
    End If
    DownloadUsingHttpRequest = False
#End If
End Function

Public Function CheckFor7Zip(Optional ByVal validate7ZipExists As Boolean = False) As Boolean
#If Mac Then
    Const LOCAL_ZIP_TOOL_FILENAME As String = "7zz"
#Else
    Const LOCAL_ZIP_TOOL_FILENAME As String = "7za.exe"
#End If

    Dim zipToolFileName As String
    Dim resourcesFolder As String
    
    If g_UserOptions.ZipSupportEnabled And Not validate7ZipExists Then
        CheckFor7Zip = True
        
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.CodeExecution.Status", INDENT_LEVEL_1, IIf(CheckFor7Zip, "Installed", "Missing"))
        End If
        
        Exit Function
    End If
    
    resourcesFolder = GetDefaultFolderPaths("Resources")
    zipToolFileName = GetZipToolFileName()
    
    If DoesFileExist(resourcesFolder & LOCAL_ZIP_TOOL_FILENAME) Then
        If IsHashValid(resourcesFolder & LOCAL_ZIP_TOOL_FILENAME, zipToolFileName) Then
            Exit Function
        End If
    End If
    
    CheckFor7Zip = Download7Zip(resourcesFolder, zipToolFileName, LOCAL_ZIP_TOOL_FILENAME)
End Function

Public Function Verify7ZipIsPresent() As Boolean
    Dim filePath As String

    filePath = GetDefaultFolderPaths("Resources") & GetZipToolFileName
    Verify7ZipIsPresent = DoesFileExist(filePath)
End Function

Public Function GetZipToolFileName() As String
#If Mac Then
    GetZipToolFileName = "7zz"
#Else
    Dim objWMI        As Object
    Dim colProcessors As Object
    Dim objProcessor  As Object

    Set objWMI = GetObject("winmgmts:\\.\root\CIMV2")
    Set colProcessors = objWMI.ExecQuery("SELECT Architecture FROM Win32_Processor")
    
    For Each objProcessor In colProcessors
        Select Case objProcessor.architecture
            Case 0: GetZipToolFileName = "7za(x86).exe"
            Case 9: GetZipToolFileName = "7za(x64).exe"
            Case 12: GetZipToolFileName = "7za(ARM).exe"
        End Select
    Next
#End If
End Function

Public Function Download7Zip(ByVal resourcesFolder As String, ByVal zipToolFileName As String, ByVal localZipToolFileName As String) As Boolean
    Dim destinationPath As String
    Dim zipFileType As String
    Dim downloadResult As Boolean
    
#If Mac Then
    zipFileType = zipToolFileName
#Else
    ' Shouldn't this then always result in 7za?
    zipFileType = Left$(zipToolFileName, Len(zipToolFileName) - 4)
#End If

    destinationPath = resourcesFolder & localZipToolFileName
    downloadResult = DownloadFile(zipToolFileName, zipFileType, destinationPath)
    
#If Mac Then
    If downloadResult Then
        Download7Zip = SetFileAsExecutable(destinationPath)
    End If
#Else
    Download7Zip = downloadResult
#End If
End Function