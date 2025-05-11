Option Explicit

#Const PRINT_DEBUG_MESSAGES = True
#If Mac Then
    Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
    Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
#End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' System and File Requirements
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function InstallFonts() As Boolean
    Dim fontURL As String
    Dim downloadEngSuccess As Boolean
    Dim downloadKorSuccess As Boolean
    
    Const FONT_BASE_URL As String = "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/"
    Const FONT_NAME_ENG As String = "just-another-hand.regular.ttf"
    Const FONT_NAME_KOR As String = "KakaoBigSans-Regular.ttf"
    Const INSTALL_FONTS_SUCCESSFUL As String = "    Font successfully installed"
    Const INSTALL_FONTS_FAILED As String = "    Unable to automatically install fonts. Please install manually."
    
    #If Mac Then
        ' No extra variables required.
    #Else
        Dim fso As Object
        Dim fontFolder As String
        Dim engFontPath As String
        Dim korFontPath As String
        Dim engSysFontPath As String
        Dim korSysFontPath As String
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Checking for Required Fonts"
    #End If
    
    #If Mac Then
        ' This is a bit hacky. THis'll be cleaned up when I have MacOS access again.
        fontURL = FONT_BASE_URL & FONT_NAME_ENG
        InstallFonts = AppleScriptTask(APPLE_SCRIPT_FILE, "InstallFonts", FONT_NAME_ENG & APPLE_SCRIPT_SPLIT_KEY & fontURL)
        If InstallFonts Then
            fontURL = FONT_BASE_URL & FONT_NAME_KOR
            InstallFonts = AppleScriptTask(APPLE_SCRIPT_FILE, "InstallFonts", FONT_NAME_KOR & APPLE_SCRIPT_SPLIT_KEY & fontURL)
        End If
    #Else
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        fontFolder = Environ$("LOCALAPPDATA") & "\Microsoft\Windows\Fonts"
        engFontPath = fso.BuildPath(fontFolder, FONT_NAME_ENG)
        korFontPath = fso.BuildPath(fontFolder, FONT_NAME_KOR)
        engSysFontPath = fso.BuildPath(Environ$("WINDIR") & "\Fonts", FONT_NAME_ENG)
        korSysFontPath = fso.BuildPath(Environ$("WINDIR") & "\Fonts", FONT_NAME_KOR)
        
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Checking Windows Paths:" & vbNewLine & _
                        "        User Font Folder: " & fontFolder & vbNewLine & _
                        "        Eng. User Path: " & engFontPath & vbNewLine & _
                        "        Kor. User Path: " & korFontPath & vbNewLine & _
                        "        Eng. System Path: " & engSysFontPath & vbNewLine & _
                        "        Kor. System Path: " & korSysFontPath
        #End If
        
        If (fso.FileExists(engFontPath) Or fso.FileExists(engSysFontPath)) And _
           (fso.FileExists(korFontPath) Or fso.FileExists(korSysFontPath)) Then
            InstallFonts = True
        Else
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "    Status: Missing" & vbNewLine & _
                            "Attempting to Install Required Fonts"
            #End If
            
            If Not DoesFolderExist(fontFolder) Then
                #If PRINT_DEBUG_MESSAGES Then
                    Debug.Print "    User font path not found. Attempting to create."
                #End If
                CreateNewFolder fontFolder
            End If
            
            If DoesFolderExist(fontFolder) Then
                #If PRINT_DEBUG_MESSAGES Then
                    Debug.Print "    Creation successful." & vbNewLine & _
                                "    Attempting to download fonts."
                #End If
                
                downloadEngSuccess = (fso.FileExists(engFontPath) Or fso.FileExists(engSysFontPath))
                downloadKorSuccess = (fso.FileExists(korFontPath) Or fso.FileExists(korSysFontPath))
                
                Select Case True
                    Case CheckForCurl()
                        If Not downloadEngSuccess Then
                            fontURL = FONT_BASE_URL & FONT_NAME_ENG
                            downloadEngSuccess = DownloadUsingCurl(engFontPath, fontURL)
                        End If
                        
                        If Not downloadKorSuccess Then
                            fontURL = FONT_BASE_URL & FONT_NAME_KOR
                            downloadKorSuccess = DownloadUsingCurl(korFontPath, fontURL)
                        End If
                    Case CheckForDotNet()
                        If Not downloadEngSuccess Then
                            fontURL = FONT_BASE_URL & FONT_NAME_ENG
                            downloadEngSuccess = DownloadUsingDotNet(engFontPath, fontURL)
                        End If
                        
                        If Not downloadKorSuccess Then
                            fontURL = FONT_BASE_URL & FONT_NAME_KOR
                            downloadKorSuccess = DownloadUsingDotNet(korFontPath, fontURL)
                        End If
                    Case Else
                        InstallFonts = False
                End Select
                
                InstallFonts = (downloadEngSuccess And downloadKorSuccess)
            Else
                #If PRINT_DEBUG_MESSAGES Then
                    Debug.Print "    Creation failed."
                #End If
                InstallFonts = False
            End If
        End If
        
        Set fso = Nothing
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        If InstallFonts Then
            Debug.Print INSTALL_FONTS_SUCCESSFUL
        Else
            Debug.Print INSTALL_FONTS_FAILED & vbNewLine & _
                        "        Just Another Hand (Regular): " & IIf(downloadEngSuccess, "Installed", "Missing") & vbNewLine & _
                        "        Kakao Big Sans (Regular):    " & IIf(downloadKorSuccess, "Installed", "Missing")
        End If
    #End If
End Function


#If Mac Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MacOS Only
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AreAppleScriptsInstalled(Optional ByVal recheckStatus As Boolean = False) As Boolean
    Dim libraryScriptsFolder As String
    Dim resourcesFolder As String
    Dim isAppleScriptInstalled As Boolean
    Dim isDialogToolkitInstalled As Boolean
    Dim statusHasBeenChecked As Boolean
    Dim scriptResult As Boolean
    
    isAppleScriptInstalled = CheckForAppleScript()
    
    If isAppleScriptInstalled Then
        If Not recheckStatus Then CheckForAppleScriptUpdate
        
        libraryScriptsFolder = "/Users/" & Environ("USER") & "/Library/Script Libraries"
        resourcesFolder = ThisWorkbook.Path & "/Resources"
        ConvertOneDriveToLocalPath resourcesFolder

        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Locating Dialog Toolkit Plus.scptd" & vbNewLine & _
                        "    Searching: " & libraryScriptsFolder
        #End If

        If Not recheckStatus Then
            ' When first opened, only check for Dialog Toolkit Plus if the folder has been previously created
            scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFolderExist", libraryScriptsFolder)
            If scriptResult Then isDialogToolkitInstalled = CheckForDialogToolkit(resourcesFolder)
        Else
            isDialogToolkitInstalled = CheckForDialogToolkit(resourcesFolder)
        End If

        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Installed: " & isDialogToolkitInstalled
        #End If

        If isDialogToolkitInstalled Then
            isDialogToolkitInstalled = CheckForDialogDisplayScript(resourcesFolder)
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "Attempting to install DialogDisplay.scpt" & vbNewLine & _
                            "    Installed: " & isDialogToolkitInstalled
            #End If
        End If
    Else
        isDialogToolkitInstalled = False
    End If

    SetVisibilityOfMacSettingsShapes isAppleScriptInstalled, isDialogToolkitInstalled

    AreAppleScriptsInstalled = isAppleScriptInstalled
End Function

Private Function AreEnhancedDialogsEnabled() As Boolean
    AreEnhancedDialogsEnabled = ThisWorkbook.Sheets("MacOS Users").Shapes("Button_EnhancedDialogs_Enable").Visible
End Function

Private Function CheckForAppleScript() As Boolean
    Dim appleScriptPath As String
    Dim appleScriptStatus As Boolean
    
    appleScriptPath = "/Users/" & Environ("USER") & "/Library/Application Scripts/com.microsoft.Excel/" & APPLE_SCRIPT_FILE
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Locating " & APPLE_SCRIPT_FILE & vbNewLine & _
                    "    Searching: " & appleScriptPath
    #End If
    
    On Error Resume Next
    appleScriptStatus = (Dir(appleScriptPath, vbDirectory) = APPLE_SCRIPT_FILE)
    On Error GoTo 0
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Found: " & appleScriptStatus
    #End If
    
    CheckForAppleScript = appleScriptStatus
End Function

Private Sub CheckForAppleScriptUpdate()
    Dim scriptFolder As String
    Dim destinationPath As String
    Dim currentScriptVersion As Long
    Dim downloadedScriptVersion As Long
    Dim appleScriptResult As Boolean
    
    Const APPLE_SCRIPT_URL As String = "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/SpeakingEvals.scpt"
    Const OLD_APPLE_SCRIPT As String = "SpeakingEvals-Old.scpt"
    Const TMP_APPLE_SCRIPT As String = "SpeakingEvals-Tmp.scpt"
    
    scriptFolder = "/Users/" & Environ("USER") & "/Library/Application Scripts/com.microsoft.Excel/"
    destinationPath = scriptFolder & TMP_APPLE_SCRIPT
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Checking if an update is available for SpeakingEvals.scpt."
    #End If
    
    On Error GoTo ErrorHandler
    
    appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DownloadFile", destinationPath & APPLE_SCRIPT_SPLIT_KEY & APPLE_SCRIPT_URL)
    If Not appleScriptResult Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Unable to download new " & APPLE_SCRIPT_FILE
        #End If
        GoTo ErrorHandler
    End If
    
    currentScriptVersion = AppleScriptTask(APPLE_SCRIPT_FILE, "GetScriptVersionNumber", "")
    downloadedScriptVersion = AppleScriptTask(TMP_APPLE_SCRIPT, "GetScriptVersionNumber", "")
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Installed Version: " & currentScriptVersion & vbNewLine & _
                    "    Online Version:    " & downloadedScriptVersion
    #End If
    
    If downloadedScriptVersion <= currentScriptVersion Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Installed version is up-to-date."
        #End If
        GoTo CleanUp
    End If
    
    appleScriptResult = AppleScriptTask(TMP_APPLE_SCRIPT, "RenameFile", scriptFolder & APPLE_SCRIPT_FILE & APPLE_SCRIPT_SPLIT_KEY & scriptFolder & OLD_APPLE_SCRIPT)
    If appleScriptResult Then appleScriptResult = AppleScriptTask(OLD_APPLE_SCRIPT, "RenameFile", scriptFolder & TMP_APPLE_SCRIPT & APPLE_SCRIPT_SPLIT_KEY & scriptFolder & APPLE_SCRIPT_FILE)
    If appleScriptResult Then appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", scriptFolder & OLD_APPLE_SCRIPT)
    If Not appleScriptResult Then GoTo ErrorHandler
    
    #If PRINT_DEBUG_MESSAGES Then
        If appleScriptResult Then Debug.Print "    Update complete."
    #End If
    
CleanUp:
    #If PRINT_DEBUG_MESSAGES Then
        If appleScriptResult Then Debug.Print "    Beginning clean up process."
    #End If
    
    On Error Resume Next
    appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", scriptFolder & TMP_APPLE_SCRIPT)
    If appleScriptResult Then
        appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", scriptFolder & TMP_APPLE_SCRIPT)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Removing temporary update file: " & IIf(appleScriptResult, "Successful", "Failed")
        #End If
    End If
    
    appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", scriptFolder & OLD_APPLE_SCRIPT)
    If appleScriptResult Then
        appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", scriptFolder & OLD_APPLE_SCRIPT)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Removing old version: " & IIf(appleScriptResult, "Successful", "Failed")
        #End If
    End If
    On Error GoTo 0
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Finished clean up."
    #End If
    Exit Sub
    
ErrorHandler:
    #If PRINT_DEBUG_MESSAGES Then
        If Err.Number <> 0 Then Debug.Print "Error during the update process."
        If Err.Description <> "" Then Debug.Print "Error: " & Err.Description
    #End If
    Resume CleanUp
End Sub

Private Function CheckForDialogToolkit(ByVal resourcesFolder As String) As Boolean
    Dim scriptResult As Boolean
    Dim libraryScriptsPath As String
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Checking for presence of Dialog Toolkit Plus." & vbNewLine & _
                    "    Local resources: " & resourcesFolder
    #End If
    
    libraryScriptsPath = AppleScriptTask(APPLE_SCRIPT_FILE, "CheckForScriptLibrariesFolder", "paramString")
    If libraryScriptsPath <> "" Then scriptResult = RequestFileAndFolderAccess(resourcesFolder, libraryScriptsPath)
    If scriptResult Then scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "InstallDialogToolkitPlus", resourcesFolder)
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Toolkit Status: " & scriptResult
    #End If
    
    CheckForDialogToolkit = scriptResult
End Function

Private Function CheckForDialogDisplayScript(ByVal resourcesFolder As String) As Boolean
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Checking for presence of DialogDisplay.scpt."
    #End If
        
    CheckForDialogDisplayScript = AppleScriptTask(APPLE_SCRIPT_FILE, "InstallDialogDisplayScript", resourcesFolder)
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Status: " & CheckForDialogDisplayScript
    #End If
End Function

Private Sub RemoveDialogToolKit(ByVal resourcesFolder As String)
    Dim scriptResult As Boolean
        
    If CheckForAppleScript() Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Removing Dialog ToolKit Plus from ~/Library/Script Libraries" & vbNewLine & _
                        "    A local copy will be stored in: " & resourcesFolder
        #End If
            
        scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "UninstallDialogToolkitPlus", resourcesFolder)
            
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Result: " & scriptResult
        #End If
    End If
End Sub

Private Sub RemindUserToInstallSpeakingEvalsScpt()
    Dim msgresult As Long
    
    Const APPLE_SCRIPT_REMINDER As String = "SpeakingEvals.scpt must be installed in order to generate reports. Please run the terminal " & _
                                            "command on the ""MacOs Users"" sheet to install it and try again."

    msgresult = DisplayMessage(APPLE_SCRIPT_REMINDER, vbOKOnly + vbExclamation, "Invalid Selection!")
    MacOS_Users.Activate
End Sub

#Else
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Windows Only
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckForCurl() As Boolean
    Dim objShell As Object
    Dim objExec As Object
    Dim checkResult As Boolean
    Dim output As String
    
    On Error GoTo ErrorHandler
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Checking if curl.exe is installed."
    #End If
    
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("cmd /c curl.exe --version")
    
    If Not objExec Is Nothing Then
        Do While Not objExec.StdOut.AtEndOfStream
            output = output & objExec.StdOut.ReadLine() & vbNewLine
        Loop
        checkResult = ((InStr(output, "curl")) > 0)
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print IIf(checkResult, "    Installed", "    Not installed. Falling back to .Net")
    #End If
    
    CheckForCurl = checkResult
CleanUp:
    If Not objExec Is Nothing Then Set objExec = Nothing
    If Not objShell Is Nothing Then Set objShell = Nothing
    Exit Function
ErrorHandler:
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Error while checking for curl.exe: " & Err.Description
    #End If
    CheckForCurl = False
    Resume CleanUp
End Function

Private Function CheckForDotNet() As Boolean
    Dim frameworkPath As String
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Verifying that Microsoft DotNet 3.5 is installed."
    #End If
    
    On Error GoTo ErrorHandler
    frameworkPath = Environ$("systemroot") & "\Microsoft.NET\Framework\v3.5"
    CheckForDotNet = Dir$(frameworkPath, vbDirectory) <> vbNullString
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "   Checking path: " & frameworkPath & vbNewLine & _
                    "   Installed: " & CheckForDotNet
    #End If
    
    Exit Function
ErrorHandler:
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Error while checking for .NET 3.5: " & Err.Description
    #End If
    CheckForDotNet = False
End Function
#End If

