Option Explicit

Private Sub Workbook_Open()
#If Mac Then
    Dim scriptResult    As Boolean
    Dim librariesFolder As String
#End If
    Dim jsonFolder      As String
    Dim resourcesFolder As String
    Dim startTime       As Date
    
    GetRunTime "Start", startTime
    
#If Mac Then
    librariesFolder = GetDefaultFolderPaths("Libraries")
#End If
    jsonFolder = GetDefaultFolderPaths("JSON")
    resourcesFolder = GetDefaultFolderPaths("Resources")
    
    InitializeDictionaries jsonFolder, True
    SetDefaultSheetVisibility
    ReadOptionsValues
    ToggleApplicationFeatures "Disable" ', True ' Triggers a log clear
    CleanUpOldLogs GetDefaultFolderPaths("Logs")

    #If Mac Then
        TriggerSystemAccessPrompt
        scriptResult = AreAppleScriptsInstalled(resourcesFolder, librariesFolder)
        If scriptResult Then
            On Error Resume Next
            AppleScriptTask APPLE_SCRIPT_FILE, "TriggerPermission", vbNullString
            On Error GoTo 0
        End If
    #End If

    ValidateOptionsValues True

    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.CodeExecution.EntryPoint", "Workbook Open"), True
    End If

    If IsFileLoadedFromTempDir Then
        #If Mac Then
            DisplayMessage "Display.Workbook.LoadedFromTempMac"
        #Else
            DisplayMessage "Display.Workbook.LoadedFromTempWindows"
        #End If

        Exit Sub
    End If
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.Workbook.BeginStartupSelfChecks", Format$(startTime, "hh:mm:ss"))
    End If
    
    If Not VerifyKeySheetsExist Then
        DisplayMessage "Display.Workbook.MissingSheets"
        GoTo ReenableEvents
    End If

    VerifySheetNames

    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.Workbook.ValidateLayouts")
    End If

    If Not ValidateSheetLayoutsOnLoad Then
        ' Display error message
        GoTo ReenableEvents
    End If

#If Mac Then
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.AppleScript.StartupInstallCheck", IIf(scriptResult, "Installed", "Missing"))
    End If
    
    SetLayoutMacOSUsers
    
    If Not scriptResult Then
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.AppleScript.ReminderToInstall")
        End If
        
        MacOS_Users.Activate
        GoTo RemindUser
    End If
#End If
    
    Instructions.Activate
    SetLayoutInstructions
    Instructions.Cells.Item(1, 3).Select

ReenableEvents:
    ToggleApplicationFeatures "Enable"
    GetRunTime "End", startTime
    Exit Sub
RemindUser:
    DisplayMessage "Display.AppleScript.InstallReminder"
    Resume ReenableEvents
End Sub

Private Sub Workbook_SheetActivate(ByVal ws As Object)
    If ws.Name Like "Chart#" Then
        Exit Sub
    End If
    
    ReadOptionsValues
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.CodeExecution.EntryPointOnActivation", ws.Name), True
    End If
    
    ToggleApplicationFeatures "Disable"
    ToggleSheetProtection ws, False
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.Worksheet.ValidateLayout", ws.Name)
    End If
    
    VerifySheetNames
    
    Select Case ws.Name
        Case "Instructions"
            SetLayoutInstructions
        Case "MacOS Users"
            ws.Shapes("cURL_Command").TextFrame2.TextRange.Characters.text = GetMsg("Textbox.MacOS.CurlCommand.Text")
            SetLayoutMacOSUsers
        Case "Options"
            SetLayoutOptions
            OptionsShapeVisibility ws
        Case Else
            If ws.Cells(1, 1).Value = "Native Teacher:" Then
                SetLayoutClassRecordsButtons ws
                ws.Cells(8, 2).Select
            End If
    End Select
    
    ToggleSheetProtection ws, True
    ToggleApplicationFeatures "Enable"
End Sub

Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)
#If Mac Then
    RemoveDialogToolKit ConvertToLocalPath(ThisWorkbook.Path & Application.PathSeparator & "Resources")
#End If
End Sub