Option Explicit

#Const Windows = (Mac = 0)

Public Function VerifyFontInstallation(Optional ByVal checkOnOpening As Boolean = False) As Boolean
    Const FONT_ENG_JAN As String = "just-another-hand"
    Const FONT_ENG_HUR As String = "Hurricane-Regular"
    Const FONT_KOR_KBS As String = "KakaoBigSans-Regular"
    
    Dim fontList() As Variant
    Dim fontInstallStatus As Variant
    Dim installationStatus As Boolean
    Dim i As Long

    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.SystemRequirements.FontsCheck")
    End If
    
    If g_UserOptions.AllFontsAreInstalled And (g_UserOptions.ValidFileHashes Or Not checkOnOpening) Then
        VerifyFontInstallation = True
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.CodeExecution.Status", vbTab, IIf(VerifyFontInstallation, "Installed", "Missing"))
        End If
        
        Exit Function
    End If
    
    fontList() = Array(FONT_ENG_JAN, FONT_ENG_HUR, FONT_KOR_KBS)
    fontInstallStatus = Array(False, False, False)
    
    For i = LBound(fontList) To UBound(fontList)
        fontInstallStatus(i) = IsFontInstalled(fontList(i))
        
        If Not fontInstallStatus(i) Then
            fontInstallStatus(i) = InstallFont(fontList(i))
        End If
    Next i
    
    installationStatus = True
    For i = LBound(fontInstallStatus) To UBound(fontInstallStatus)
        If fontInstallStatus(i) = False Then
            installationStatus = False
            Exit For
        End If
    Next i
    
    With Options
        ToggleSheetProtection Options, False
        WriteNewRangeValue .Range(g_FONT_INSTALLATION_STATUS), IIf(installationStatus, "Yes", "No")
        WriteNewRangeValue .Range(g_VALID_HASHES_STATUS), IIf(installationStatus, "Yes", "No")
        ToggleSheetProtection Options, True
    End With
    
    If g_UserOptions.EnableLogging Then
        If VerifyFontInstallation Then
            DebugAndLogging GetMsg("Debug.SystemRequirements.FontsInstallationSuccessful")
        Else
            DebugAndLogging GetMsg("Debug.SystemRequirements.FontsInstallationFailed", IIf(fontInstallStatus(0), "Installed", "Missing"), IIf(fontInstallStatus(1), "Installed", "Missing"), IIf(fontInstallStatus(2), "Installed", "Missing"))
        End If
    End If
    
    VerifyFontInstallation = installationStatus
End Function

Private Function IsFontInstalled(ByVal fontName As String) As Boolean
    Dim fontFileName As String
    Dim fontInstalled As Boolean
    
    fontFileName = ReadValueFromDictionary(g_dictFileData, fontName, "filename")

#If Mac Then
    fontInstalled = AppleScriptTask(APPLE_SCRIPT_FILE, "IsFontInstalled", fontFileName)
#Else
    Dim fso             As Object
    Dim localFontPath   As String
    Dim sysfontPath     As String
    Dim fontPath        As String
    Dim fontDisplayName As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    localFontPath = GetDefaultFolderPaths("Windows Local Font Path") & fontFileName
    sysfontPath = GetDefaultFolderPaths("Windows System Font Path") & fontFileName
    fontDisplayName = ReadValueFromDictionary(g_dictFileData, fontName, "englishdisplayname")

    If fso.fileExists(localFontPath) Then
        fontInstalled = True
        fontPath = localFontPath
    ElseIf fso.fileExists(sysfontPath) Then
        fontInstalled = True
        fontPath = sysfontPath
    End If

    If fontPath <> vbNullString Then
        If Not IsFontRegistered(fontDisplayName) Then
            fontInstalled = RegisterFont(fontPath, fontDisplayName)
        End If
    End If
#End If

    IsFontInstalled = fontInstalled
End Function

Private Function InstallFont(ByVal fontName As String) As Boolean
    Dim fontFileName As String

    fontFileName = ReadValueFromDictionary(g_dictFileData, fontName, "filename")

#If Mac Then
    Dim fontURL As String
    
    fontURL = ReadValueFromDictionary(g_dictFileData, fontName, "url")

    On Error Resume Next
    InstallFont = AppleScriptTask(APPLE_SCRIPT_FILE, "InstallFonts", fontFileName & APPLE_SCRIPT_SPLIT_KEY & fontURL)
    On Error GoTo 0
#Else
    Dim downloadFilePath  As String
    Dim localFontFolder   As String
    Dim fontDisplayName   As String
    Dim fontInstalled     As Boolean
    
    downloadFilePath = GetDefaultFolderPaths("Resources") & fontFileName
    localFontFolder = GetDefaultFolderPaths("Windows Local Font Path")
    fontDisplayName = ReadValueFromDictionary(g_dictFileData, fontName, "englishdisplayname")
    
    If CheckForAndAttemptToCreateFolder(localFontFolder) Then
        If Not DoesFileExist(downloadFilePath) Then
            If Not DownloadFile(fontName, fontName, downloadFilePath) Then
                InstallFont = False
                Exit Function
            End If
        End If
        
        fontInstalled = InstallFontLocally(downloadFilePath, localFontFolder, fontFileName, fontDisplayName)
    End If

    InstallFont = fontInstalled
#End If
End Function

#If Windows Then
Private Function InstallFontLocally(ByVal downloadedFontPath As String, ByVal localFontFolder As String, ByVal fontFileName As String, ByVal fontDisplayName As String) As Boolean
    Dim installSuccessful As Boolean
    
    If MoveFile(downloadedFontPath, localFontFolder) Then
        installSuccessful = RegisterFont(localFontFolder & fontFileName, fontDisplayName)
    End If
    
    InstallFontLocally = installSuccessful
End Function

Private Function RegisterFont(ByVal fontPath As String, ByVal fontDisplayName As String) As Boolean
    Dim WshShell As Object: Set WshShell = CreateObject("WScript.Shell")
    
    On Error Resume Next
    WshShell.RegWrite "HKCU\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts\" & fontDisplayName & " (TrueType)", fontPath, "REG_SZ"
    On Error GoTo 0

    RegisterFont = (Err.Number = 0)
End Function

Private Function IsFontRegistered(ByVal fontDisplayName As String) As Boolean
    Dim WshShell As Object: Set WshShell = CreateObject("WScript.Shell")
    Dim fontFile As String
    
    On Error Resume Next
    fontFile = WshShell.RegRead("HKCU\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts\" & fontDisplayName & " (TrueType)")
    On Error GoTo 0

    IsFontRegistered = (Err.Number = 0)
End Function
#End If