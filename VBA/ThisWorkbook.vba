Option Explicit

#Const PRINT_DEBUG_MESSAGES = True
#If Mac Then
    Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
    Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
#End If

Private Sub Workbook_Open()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim shps As Shapes
    Dim startTime As Date
    Dim endTime As Date
    Dim elapsedTime As Double
    Dim msgResult As Long
    
    #If Mac Then
        Dim scriptResult As Boolean
    #End If
    
    Const CURL_COMMAND_TEXT As String = "curl -L -o ~/Library/Application\ Scripts/com.microsoft.Excel/SpeakingEvals.scpt https://github.com/papercutter0324/SpeakingEvals/raw/main/AppleScript/SpeakingEvals.scpt"
    Const STARTUP_MSG_TEMP_DIR As String = "This file has been loaded from a temporary folder and will not function correctly. " & _
                                           "Please verify you have correctly extracted this file from the zip file (if applicable) " & _
                                           "and save it to a permanent location."
    Const STARTUP_MSG_COMPLETE As String = "Self-Check complete!"
    Const STARTUP_MSG_APPLE_SCRIPT_REMINDER As String = "You must install SpeakingEvals.scpt for this file to fuction properly. Please follow the " & _
                                                        "installation instructions and read the notices about the System Events and File & Folder " & _
                                                        "Permission requests."
    Const STARTUP_MSG_TEMP_DIR_SIZE As Long = 470
    Const STARTUP_MSG_INITIAL_SIZE As Long = 430
    Const STARTUP_MSG_COMPLETE_SIZE As Long = 180
    Const STARTUP_MSG_APPLE_SCRIPT_REMINDER_SIZE As Long = 470
    
    ' This might need to be moved elsewhere to ensure it is diplayed properly
    If IsFileLoadedFromTempDir Then
        msgResult = DisplayMessage(STARTUP_MSG_TEMP_DIR, vbOKOnly + vbExclamation, "Warning!", STARTUP_MSG_TEMP_DIR_SIZE)
        Exit Sub
    End If
    
    startTime = Now
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Beginning start-up self-checks" & vbNewLine & _
                    INDENT_LEVEL_1 & "Start Time: " & Format$(startTime, "hh:mm:ss")
    #End If
    
    ToggleApplicationFeatures False
    
    VerifySheetNames
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Validating Layouts"
    #End If
    
    On Error GoTo ReenableEvents
    Set wb = ThisWorkbook
    For Each ws In wb.Worksheets
        ToggleSheetProtection ws, False
        
        With ws
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print INDENT_LEVEL_1 & "Sheet: " & .Name
            #End If
            
            Select Case .Name
                Case "Instructions"
                    If .Visible = xlSheetHidden Then .Visible = xlSheetVisible
                Case "MacOS Users"
                    #If Mac Then
                        If .Visible = xlSheetHidden Then .Visible = xlSheetVisible
                        
                        Set shps = .Shapes
                        shps.[_Default]("cURL_Command").TextFrame2.TextRange.Characters.Text = CURL_COMMAND_TEXT
                        scriptResult = AreAppleScriptsInstalled()
                        Set shps = Nothing
                    #Else
                        If .Visible <> xlSheetHidden Then .Visible = xlSheetHidden
                    #End If
                Case "Options"
                    If .Visible = xlSheetHidden Then .Visible = xlSheetVisible
                    SetLayoutOptions
                Case Else
                    If .Visible = xlSheetHidden Then .Visible = xlSheetVisible
                    AutoPopulateEvaluationDateValues ws
                    SetLayoutClassRecords ws
            End Select
        End With
        
        ToggleSheetProtection ws, True
    Next ws
    
    ' Calculate time now to avoid user response time being added to the result
    endTime = Now
    elapsedTime = endTime - startTime
    
    #If Mac Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "SpeakingEvals.scpt" & vbNewLine & _
                        INDENT_LEVEL_1 & "Status: " & IIf(scriptResult, "Installed", "Missing")
        #End If
        If Not scriptResult Then
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print INDENT_LEVEL_1 & "Remindering user to install." & vbNewLine & _
                            INDENT_LEVEL_1 & "Activating sheet ""MacOS Users"""
            #End If
            MacOS_Users.Activate
            SetLayoutMacOSUsers
            msgResult = DisplayMessage(STARTUP_MSG_APPLE_SCRIPT_REMINDER, vbOKOnly + vbInformation, "Notice!", STARTUP_MSG_APPLE_SCRIPT_REMINDER_SIZE)
        Else
            Instructions.Activate
            SetLayoutInstructions
            Instructions.Cells.Item(1, 3).Select
        End If
    #Else
        Instructions.Activate
        SetLayoutInstructions
        Instructions.Cells.Item(1, 3).Select
    #End If

ReenableEvents:
    ToggleApplicationFeatures True
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Finished Tasks" & vbNewLine & _
                    INDENT_LEVEL_1 & "End Time: " & Format$(endTime, "hh:mm:ss") & vbNewLine & _
                    INDENT_LEVEL_1 & "Elapsed time: " & Format$(elapsedTime * 86400, "0.00") & " seconds"
    #End If
End Sub

Private Sub Workbook_SheetActivate(ByVal ws As Object)
    Const CURL_COMMAND_TEXT As String = "curl -L -o ~/Library/Application\ Scripts/com.microsoft.Excel/SpeakingEvals.scpt https://github.com/papercutter0324/SpeakingEvals/raw/main/AppleScript/SpeakingEvals.scpt"
    
    ToggleApplicationFeatures False
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Validating Layout" & vbNewLine & _
                    INDENT_LEVEL_1 & "Sheet: " & ws.Name
    #End If
    
    VerifySheetNames
    
    Select Case ws.Name
        Case "Instructions"
            SetLayoutInstructions
        Case "MacOS Users"
            ws.Shapes("cURL_Command").TextFrame2.TextRange.Characters.Text = CURL_COMMAND_TEXT
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
    
    ToggleApplicationFeatures True
End Sub

Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)
    #If Mac Then
        Dim resourcesFolder As String
        
        resourcesFolder = ConvertOneDriveToLocalPath(ThisWorkbook.Path & Application.PathSeparator & "Resources")
        RemoveDialogToolKit resourcesFolder
    #End If
End Sub
