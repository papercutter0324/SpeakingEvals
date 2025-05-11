Option Explicit

#Const PRINT_DEBUG_MESSAGES = True
#If Mac Then
    Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
    Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
#End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Message Display
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DisplayWarning(ByVal msgTitle As String, Optional ByVal inputLength As Long = 0)
    Dim msgToDisplay As String
    Dim msgDialogType As Long
    Dim msgDialogWidth As Long
    Dim msgresult As Long
    
    Select Case msgTitle
        Case "English Name: Exceeds Max Length"
            msgToDisplay = "The student's English name is longer than 30 characters and may not " & _
                           "fit on the report. Please verify how it looks after generating " & _
                           "the report and consider using a shorter version." & vbNewLine & vbNewLine & _
                           "Report generation will still work."
            msgDialogType = vbOKOnly + vbInformation
            msgDialogWidth = 370
        Case "Korean Name: Uncommon Length"
            msgToDisplay = "You entered a Korean name with " & CStr(inputLength) & " syllables. These names do exist, " & _
                           "but they are uncommon. Please verify you have typed it correctly and using Hangul." & vbNewLine & vbNewLine & _
                           "Report generation will still work."
            msgDialogType = vbOKOnly + vbInformation
            msgDialogWidth = 380
        Case "Korean Name: Invalid Length"
            msgToDisplay = "You entered an invalid name length. Please verify you have typed it correctly and using Hangul."
            msgDialogType = vbOKOnly + vbExclamation
            msgDialogWidth = 380
        Case "Date: Invalid Format"
            msgToDisplay = "Please enter a valid date."
            msgDialogType = vbOKOnly + vbExclamation
            msgDialogWidth = 200
        Case "Grade: Invalid Score"
            msgToDisplay = "An invalid score value has been entered. Please enter A+, A, B+, B, C, or a number between 1 and 5."
            msgDialogType = vbOKOnly + vbExclamation
            msgDialogWidth = 250
        Case "Comment: Too Short"
            msgToDisplay = "The comment you have typed is very short (under 80 characters). Please check that you " & _
                           "have followed the ""Positive - Negative - Positive"" format and provided sufficient detail."
            msgDialogType = vbOKOnly + vbInformation
            msgDialogWidth = 280
        Case "Comment: Exceeds Max Length"
            msgToDisplay = "The comment you have typed is too long (" & CStr(inputLength) & " chars). Please shorten it by at least " & _
                           Len(inputLength) - 960 & " characters to ensure it fits in the report's comment box."
            msgDialogType = vbOKOnly + vbExclamation
            msgDialogWidth = 300
    End Select
    
    msgresult = DisplayMessage(msgToDisplay, msgDialogType, msgTitle, msgDialogWidth)
End Sub

Public Function DisplayMessage(ByVal messageText As String, ByVal messageType As Long, ByVal messageTitle As String, Optional ByVal dialogWidth As Long = 250) As Variant
    #If Mac Then
        Dim dialogType As String
        Dim iconType As String
        Dim dialogParameters As String
        Dim dialogResult As Variant
        Dim messageDisplayed As Boolean
        Dim lastError As String
        Dim i As Long
        
        ' Button types for bitwise comparison
        Const BUTTON_OK_ONLY As Long = 0
        Const BUTTON_OK_CANCEL As Long = 1
        Const BUTTON_RETRY_CANCEL As Long = 2
        Const BUTTON_YES_NO As Long = 4
        Const BUTTON_YES_NO_CANCEL As Long = 8
        
        ' Icon types for bitwise comparison
        Const ICON_CRITICAL As Long = 16
        Const ICON_QUESTION As Long = 32
        Const ICON_EXCLAMATION As Long = 48
        Const ICON_INFORMATION As Long = 64

        If AreEnhancedDialogsEnabled Then
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "Attempting to display message via Dialog Toolkit Plus." & vbNewLine & _
                            "    Message: " & messageText
            #End If
            
            ' Determine buttons to display
            Select Case True
                Case (messageType And BUTTON_OK_ONLY) = BUTTON_OK_ONLY
                    dialogType = "OkOnly"
                Case (messageType And BUTTON_OK_CANCEL) = BUTTON_OK_CANCEL
                    dialogType = "OkCancel"
                Case (messageType And BUTTON_RETRY_CANCEL) = BUTTON_RETRY_CANCEL
                    dialogType = "RetryCancel"
                Case (messageType And BUTTON_YES_NO) = BUTTON_YES_NO
                    dialogType = "YesNo"
                Case (messageType And BUTTON_YES_NO_CANCEL) = BUTTON_YES_NO_CANCEL
                    dialogType = "YesNoCancel"
                Case Else
                    dialogType = "OkOnly"
            End Select
            
            ' Determine icon to display
            Select Case True
                Case (messageType And ICON_CRITICAL) = ICON_CRITICAL
                    iconType = "CriticalIcon"
                Case (messageType And ICON_QUESTION) = ICON_QUESTION
                    iconType = "QuestionIcon"
                Case (messageType And ICON_EXCLAMATION) = ICON_EXCLAMATION
                    iconType = "ExclamationIcon"
                Case (messageType And ICON_INFORMATION) = ICON_INFORMATION
                    iconType = "InformationIcon"
                Case Else
                    iconType = "OtherIcon"
            End Select
                    
            ' Update SpeakingEvals.scpt to supoport:
            ' dialogParameters = messageText & APPLE_SCRIPT_SPLIT_KEY & dialogType & APPLE_SCRIPT_SPLIT_KEY & iconType & APPLE_SCRIPT_SPLIT_KEY & messageTitle & APPLE_SCRIPT_SPLIT_KEY & dialogWidth
            dialogParameters = messageText & APPLE_SCRIPT_SPLIT_KEY & dialogType & APPLE_SCRIPT_SPLIT_KEY & messageTitle & APPLE_SCRIPT_SPLIT_KEY & dialogWidth
            
            On Error Resume Next
            Do While Not messageDisplayed
                dialogResult = AppleScriptTask("DialogDisplay.scpt", "DisplayDialog", dialogParameters)
                messageDisplayed = (dialogResult <> "")
                i = i + 1
                
                #If PRINT_DEBUG_MESSAGES Then
                    If Err.Number <> 0 Then lastError = Err.Number & " - " & Err.Description
                #End If
                
                If i >= 10 Then
                    dialogResult = MsgBox(messageText, messageType, messageTitle)
                    messageDisplayed = True
                End If
            Loop
            On Error GoTo 0
            
            #If PRINT_DEBUG_MESSAGES Then
                If lastError = "" Then lastError = "N/A"
                Debug.Print "    Number of attempts: " & i & vbNewLine & _
                            "    Final error: " & lastError
            #End If
            
            DisplayMessage = dialogResult
            Exit Function
        End If
    #End If
    
    DisplayMessage = MsgBox(messageText, messageType, messageTitle)
End Function
