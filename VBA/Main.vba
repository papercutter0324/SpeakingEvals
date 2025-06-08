Option Explicit

#Const PRINT_DEBUG_MESSAGES = True
#If Mac Then
    Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
    Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
#End If

Public Const INDENT_LEVEL_1 As String = "    "
Public Const INDENT_LEVEL_2 As String = INDENT_LEVEL_1 & INDENT_LEVEL_1
Public Const INDENT_LEVEL_3 As String = INDENT_LEVEL_1 & INDENT_LEVEL_2

Public Const RANGE_NATIVE_TEACHER As String = "C1"
Public Const RANGE_KOREAN_TEACHER As String = "C2"
Public Const RANGE_CLASS_LEVEL As String = "C3"
Public Const RANGE_CLASS_DAYS As String = "C4"
Public Const RANGE_CLASS_TIME As String = "C5"
Public Const RANGE_EVAL_DATE As String = "C6"
Public Const RANGE_ENGLISH_NAME As String = "B8:B32"
Public Const RANGE_KOREAN_NAME As String = "C8:C32"
Public Const RANGE_FULL_NAME As String = "B8:C32"
Public Const RANGE_GRADES As String = "D8:I32"
Public Const RANGE_COMMENT As String = "J8:J32"
Public Const RANGE_NOTES As String = "K8:M32"
Public Const RANGE_WINNERS As String = "L2:L4"
Public Const RANGE_VALIDATION_LIST As String = "O8:O32"
Public Const RANGE_ALL_MONITORED As String = "C1:C2,C6,B8:J32,L2:L4"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Main() and Shared Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Main()
    Dim ws As Worksheet
    Dim clickedButtonName As String
    Dim msgResult As Long
    Dim startTime As Date
    Dim endTime As Date
    Dim elapsedTime As Double
    
    #If Mac Then
        Dim msgToDiplay As String
    #End If
    
    ' Add a static variable or track a hidden cell so that this only appears once per instance
    ' Update this to mention an occasional PPT hang when closing.
    Const STARTUP_MSG_TEMP_DIR As String = "This file has been loaded from a temporary folder and will not function correctly. " & _
                                           "Please verify you have correctly extracted this file from the zip file (if applicable) " & _
                                           "and save it to a permanent location."
    Const STARTUP_MSG_TEMP_DIR_SIZE As Long = 470
    
    startTime = Now
    
    Set ws = ActiveSheet
    clickedButtonName = Application.Caller
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Beginning Tasks" & vbNewLine & _
                    INDENT_LEVEL_1 & "Start Time: " & Format$(startTime, "hh:mm:ss") & vbNewLine & _
                    INDENT_LEVEL_1 & "Active Worksheet: " & ws.Name & vbNewLine & _
                    INDENT_LEVEL_1 & "Button Pressed: " & clickedButtonName
    #End If
    
    ToggleApplicationFeatures False
    
    If IsFileLoadedFromTempDir Then
        msgResult = DisplayMessage(STARTUP_MSG_TEMP_DIR, vbOKOnly + vbExclamation, "Warning!", STARTUP_MSG_TEMP_DIR_SIZE)
        GoTo ReenableEvents
    End If
    
    On Error GoTo ReenableEvents
    Select Case clickedButtonName
        Case "Button_SignatureEmbedded", "Button_SignatureMissing"
            ToggleEmbeddedSignature clickedButtonName
        Case "Button_RepairLayout"
            RepairLayouts ws
        Case "Button_AutoSelectWinners"
            AutoSelectClassWinners ws
        Case "Button_GenerateCertificates"
            'msgResult = DisplayMessage( _
                "This feature has not been implemented yet, but it is planned for an upcoming " & _
                "update. Sorry for the inconvenience.", vbOKOnly + vbInformation, "Notice!")
            'GoTo ReenableEvents
    End Select
    
    #If Mac Then
        If Not AreAppleScriptsInstalled(True) Then
            RemindUserToInstallSpeakingEvalsScpt
            GoTo ReenableEvents
        End If
        
        Select Case clickedButtonName
            Case "Button_EnhancedDialogs_Enable", "Button_EnhancedDialogs_Disable"
                ToogleMacSettingsButtons ws, clickedButtonName
        End Select
    #End If
    
    Select Case clickedButtonName
        Case "Button_GenerateReports", "Button_GenerateProofs", "Button_GenerateCertificates"
            msgResult = DisplayMessage( _
                "There is a uncommon error, where the first time you try to save it fails. " & _
                "If you experience this, wait a couple seconds and try again. It should work " & _
                "fine the second time." & vbNewLine & vbNewLine & "Press okay to continue " & _
                "creating the reports", vbOKOnly + vbInformation, "Notice!")
                
            ' GenerateReports ws, clickedButtonName
            CreateReportsAndCertificates ws, clickedButtonName
            ws.Activate
    End Select
    
ReenableEvents:
    endTime = Now
    elapsedTime = endTime - startTime
    
    ToggleApplicationFeatures True
        
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Finished Tasks" & vbNewLine & _
                    INDENT_LEVEL_1 & "End Time: " & Format$(endTime, "hh:mm:ss") & vbNewLine & _
                    INDENT_LEVEL_1 & "Elapsed time: " & Format$(elapsedTime * 86400, "0.00") & " seconds" & vbNewLine & vbNewLine
    #End If
End Sub

Public Sub ToggleSheetProtection(ByVal ws As Worksheet, ByVal protectionStatus As Boolean)
    With ws
        If protectionStatus Then
            .Protect
            .EnableSelection = xlUnlockedCells
        Else
            .Unprotect
        End If
    End With
End Sub

Public Sub ToggleApplicationFeatures(ByVal enabledStatues As Boolean)
    With Application
        .Calculation = IIf(enabledStatues, xlCalculationAutomatic, xlCalculationManual)
        .EnableAnimations = enabledStatues
        .EnableEvents = enabledStatues
        .ScreenUpdating = enabledStatues
    
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Toggling Application Updates" & vbNewLine & _
                        INDENT_LEVEL_1 & "Calculation: " & IIf(.Calculation = xlCalculationManual, "Manual", "Automatic") & vbNewLine & _
                        INDENT_LEVEL_1 & "EnableEvents: " & .EnableEvents & vbNewLine & _
                        INDENT_LEVEL_1 & "ScreenUpdating: " & .ScreenUpdating
        #End If
    End With
End Sub

Public Function GetCellType(ByVal changedCell As Range) As String
    With changedCell.Worksheet
        Select Case True
            Case Not Intersect(changedCell, .Range(RANGE_NATIVE_TEACHER)) Is Nothing
                GetCellType = "Native Teacher"
            Case Not Intersect(changedCell, .Range(RANGE_KOREAN_TEACHER)) Is Nothing
                GetCellType = "Korean Teacher"
            Case Not Intersect(changedCell, .Range(RANGE_CLASS_LEVEL)) Is Nothing
                GetCellType = "Level"
            Case Not Intersect(changedCell, .Range(RANGE_CLASS_DAYS)) Is Nothing
                GetCellType = "Class Days"
            Case Not Intersect(changedCell, .Range(RANGE_CLASS_TIME)) Is Nothing
                GetCellType = "Class Time"
            Case Not Intersect(changedCell, .Range(RANGE_EVAL_DATE)) Is Nothing
                GetCellType = "Eval Date"
            Case Not Intersect(changedCell, .Range(RANGE_ENGLISH_NAME)) Is Nothing
                GetCellType = "English Name"
            Case Not Intersect(changedCell, .Range(RANGE_KOREAN_NAME)) Is Nothing
                GetCellType = "Korean Name"
            Case Not Intersect(changedCell, .Range(RANGE_GRADES)) Is Nothing
                GetCellType = "Grade"
            Case Not Intersect(changedCell, .Range(RANGE_COMMENT)) Is Nothing
                GetCellType = "Comment"
            Case Not Intersect(changedCell, .Range(RANGE_NOTES)) Is Nothing
                GetCellType = "Notes"
            Case Not Intersect(changedCell, .Range(RANGE_WINNERS)) Is Nothing
                GetCellType = "Winner Names"
            Case Else
                GetCellType = "Unknown"
        End Select
    End With
End Function

Public Function TrimStringBeforeCharacter(ByRef stringToTrim As String, Optional ByVal trimPoint As String = "(") As String
    Dim charPos As Long
    
    charPos = InStr(stringToTrim, trimPoint)
    If charPos > 0 Then
        stringToTrim = Left$(stringToTrim, charPos - 1)
    End If
    
    TrimStringBeforeCharacter = stringToTrim
End Function
