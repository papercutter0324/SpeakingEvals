Option Explicit

#Const PRINT_DEBUG_MESSAGES = True
#If Mac Then
    Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
    Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
#End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Message Display
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                            INDENT_LEVEL_1 & "Message: " & messageText
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
                Debug.Print INDENT_LEVEL_1 & "Number of attempts: " & i & vbNewLine & _
                            INDENT_LEVEL_1 & "Final error: " & lastError
            #End If
            
            DisplayMessage = dialogResult
            Exit Function
        End If
    #End If
    
    DisplayMessage = MsgBox(messageText, messageType, messageTitle)
End Function
