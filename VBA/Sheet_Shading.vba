Option Explicit

#Const PRINT_DEBUG_MESSAGES = True
#If Mac Then
    Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
    Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
#End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Cell and Data Shading
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetDefaultShading(ByVal ws As Worksheet)
    Dim englishNameRange As Range
    Dim koreanNameRange As Range
    Dim commentRange As Range
    Dim validationListRange As Range
    Dim currentCell As Range
    Dim shadingUpdates As New Dictionary
    Dim shadingKey As Variant
    Dim nameToFind As String
    Dim shadingValue As Long
    Dim currentRow As Long
    Dim i As Long
    
    With ws
        Set englishNameRange = .Range(RANGE_ENGLISH_NAME)
        Set koreanNameRange = .Range(RANGE_KOREAN_NAME)
        Set commentRange = .Range(RANGE_COMMENT)
        Set validationListRange = .Range(RANGE_VALIDATION_LIST)
    End With
    
    ' Set shadingUpdates = CreateObject("Scripting.Dictionary")
    
    For Each currentCell In englishNameRange
        With currentCell
            AddToShadingDictionary shadingUpdates, .Address, GetEnglishNameShading(.Value, False)
        End With
    Next currentCell
        
    For Each currentCell In koreanNameRange
        With currentCell
            AddToShadingDictionary shadingUpdates, .Address, GetKoreanNameShading(.Value, False)
        End With
    Next currentCell
    
    For Each currentCell In commentRange
        With currentCell
            AddToShadingDictionary shadingUpdates, .Address, GetCommentShading(.Value, False)
        End With
    Next currentCell
        
    For i = 2 To 4
        nameToFind = ws.Range("L" & i).Value
        shadingValue = GetWinnerShadingValue("$L$" & i)

        If nameToFind <> vbNullString Then
            SetShadingForWinnerName validationListRange, nameToFind, shadingUpdates, shadingValue
        End If
    Next i
    
    ApplyShading ws, shadingUpdates
End Sub

Public Function GetEnglishNameShading(ByVal nameValue As String, Optional ByVal enableWarningMsg As Boolean = True) As Long
    Dim msgTitle As String
    Dim msgToDisplay As String
    Dim msgDialogType As Long
    Dim msgDialogWidth As Long
    
    Select Case Len(nameValue)
        Case 0 ' Empty cell
            GetEnglishNameShading = RGB(255, 255, 255)
        Case Is <= 21 ' Within acceptable length
            GetEnglishNameShading = RGB(255, 255, 255)
        Case Else ' Too long
            If enableWarningMsg Then
                msgTitle = "English Name: Exceeds Max Length"
                msgToDisplay = "The student's English name is longer than 40 characters and may not " & _
                               "fit on the report. Please verify how it looks after generating " & _
                               "the report and consider using a shorter version." & vbNewLine & vbNewLine & _
                               "Report generation will still work."
                msgDialogType = vbOKOnly + vbInformation
                msgDialogWidth = 370
                Call DisplayMessage(msgToDisplay, msgDialogType, msgTitle, msgDialogWidth)
            End If
            GetEnglishNameShading = RGB(255, 255, 0)
    End Select
End Function

Public Function GetKoreanNameShading(ByVal nameValue As String, Optional ByVal enableWarningMsg As Boolean = True) As Long
    Dim msgTitle As String
    Dim msgToDisplay As String
    Dim msgDialogType As Long
    Dim msgDialogWidth As Long
    
    nameValue = TrimStringBeforeCharacter(nameValue)
    
    Select Case Len(nameValue)
        Case 0, 3 ' Empty cell or typical name length
            GetKoreanNameShading = RGB(255, 255, 255)
        Case 2, 4 ' Uncommon but possible
            If enableWarningMsg Then
                msgTitle = "Korean Name: Uncommon Length"
                msgToDisplay = "You entered a Korean name with " & CStr(Len(nameValue)) & " syllables. These names do exist, " & _
                           "but they are uncommon. Please verify you have typed it correctly and using Hangul." & vbNewLine & vbNewLine & _
                           "Report generation will still work."
                msgDialogType = vbOKOnly + vbInformation
                msgDialogWidth = 380
                Call DisplayMessage(msgToDisplay, msgDialogType, msgTitle, msgDialogWidth)
            End If
            GetKoreanNameShading = RGB(255, 255, 0)
        Case Else ' Includes 1 and > 4, likely errors
            If enableWarningMsg Then
                msgTitle = "Korean Name: Invalid Length"
                msgToDisplay = "You entered an invalid name length. Please verify you have typed it correctly and using Hangul."
                msgDialogType = vbOKOnly + vbExclamation
                msgDialogWidth = 380
                Call DisplayMessage(msgToDisplay, msgDialogType, msgTitle, msgDialogWidth)
            End If
            GetKoreanNameShading = RGB(255, 0, 0)
    End Select
End Function

Public Function GetDefaultGradeShading(ByVal columnNumber As Long) As Long
    Select Case columnNumber
        Case 4: GetDefaultGradeShading = RGB(228, 223, 236) ' Grammar
        Case 5: GetDefaultGradeShading = RGB(218, 238, 243) ' Pronunciation
        Case 6: GetDefaultGradeShading = RGB(242, 220, 219) ' Fluency
        Case 7: GetDefaultGradeShading = RGB(253, 233, 217) ' Manner
        Case 8: GetDefaultGradeShading = RGB(235, 241, 222) ' Content
        Case 9: GetDefaultGradeShading = RGB(220, 230, 241) ' Effort
        Case Else: GetDefaultGradeShading = xlNone          ' Default or error
    End Select
End Function

Public Function GetCommentShading(ByVal commentValue As String, Optional ByVal enableWarningMsg As Boolean = True) As Long
    Const MIN_LEN As Long = 80
    Const MAX_LEN As Long = 960
    
    Dim msgTitle As String
    Dim msgToDisplay As String
    Dim msgDialogType As Long
    Dim msgDialogWidth As Long

    Select Case Len(commentValue)
        Case 0 ' Empty cell
            GetCommentShading = RGB(242, 242, 242)
        Case 1 To MIN_LEN - 1 ' Comment is too short
            If enableWarningMsg Then
                msgTitle = "Comment: Too Short"
                msgToDisplay = "The comment you have typed is very short (under 80 characters). Please check that you " & _
                           "have followed the ""Positive - Negative - Positive"" format and provided sufficient detail."
                msgDialogType = vbOKOnly + vbInformation
                msgDialogWidth = 280
                Call DisplayMessage(msgToDisplay, msgDialogType, msgTitle, msgDialogWidth)
            End If
            GetCommentShading = RGB(255, 255, 0)
        Case Is > MAX_LEN ' Comment exceeds max length
            If enableWarningMsg Then
                msgTitle = "Comment: Exceeds Max Length"
                msgToDisplay = "The comment you have typed is too long (" & CStr(Len(commentValue)) & " chars). Please shorten it by at least " & _
                               CStr(Len(commentValue) - 960) & " characters to ensure it fits in the report's comment box."
                msgDialogType = vbOKOnly + vbExclamation
                msgDialogWidth = 300
                Call DisplayMessage(msgToDisplay, msgDialogType, msgTitle, msgDialogWidth)
            End If
            GetCommentShading = RGB(255, 0, 0)
        Case Else ' Within acceptable length
            GetCommentShading = RGB(242, 242, 242)
    End Select
End Function

Public Function GetWinnerShadingValue(ByVal placementCellAddress As String) As Long
    Dim shadingValue As Long
    
    Select Case placementCellAddress
        Case "$L$2"
            shadingValue = RGB(255, 215, 0)
        Case "$L$3"
            shadingValue = RGB(192, 192, 192)
        Case "$L$4"
            shadingValue = RGB(205, 127, 50)
    End Select
    
    GetWinnerShadingValue = shadingValue
End Function

Public Sub SetShadingForWinnerName(ByVal validationListRange As Range, ByVal enteredValue As String, ByRef shadingUpdates As Dictionary, ByVal shadingValue As Long)
    Dim cellToQuery As Range
    
    For Each cellToQuery In validationListRange
        If cellToQuery.Value = enteredValue Then
             AddToShadingDictionary shadingUpdates, "$B$" & cellToQuery.Row & ":$C$" & cellToQuery.Row, shadingValue
             Exit For
         End If
    Next cellToQuery
End Sub

Public Sub DetermineNameCellShading(ByVal ws As Worksheet, ByVal cellRange As Range, ByVal shadingValue As Long, ByRef shadingUpdates As Dictionary)
    Dim cellToQuery As Range
    Dim currentRow As Long
    
    For Each cellToQuery In cellRange
        If cellToQuery.Interior.Color = shadingValue Then
            currentRow = cellToQuery.Row
            AddToShadingDictionary shadingUpdates, "$B$" & currentRow, GetEnglishNameShading(ws.Range("B" & currentRow).Value, False)
            AddToShadingDictionary shadingUpdates, "$C$" & currentRow, GetKoreanNameShading(ws.Range("C" & currentRow).Value, False)
            Exit For
        End If
    Next cellToQuery
End Sub

Public Sub RemoveDuplicateWinners(ByVal ws As Worksheet, ByVal cellRange As Range, ByVal currentCellAddress As String, ByVal shadingValue As Long, ByRef shadingUpdates As Dictionary, ByVal enteredValue As String)
    Dim cellToQuery As Range
    Dim currentRow As Long

    For Each cellToQuery In cellRange
        With cellToQuery
            If .Address <> currentCellAddress And .Value = enteredValue Then
                .Value = vbNullString
                currentRow = .Row
                If currentRow >= 8 And currentRow <= 32 Then
                    AddToShadingDictionary shadingUpdates, "$B$" & currentRow, GetEnglishNameShading(ws.Range("B" & currentRow).Value, False)
                    AddToShadingDictionary shadingUpdates, "$C$" & currentRow, GetKoreanNameShading(ws.Range("C" & currentRow).Value, False)
                End If
            End If
        End With
    Next cellToQuery
End Sub

Public Sub AddToShadingDictionary(ByRef shadingDictionary As Object, ByVal cellToShade As String, ByVal shadingValue As Long)
    If Not shadingDictionary.Exists(cellToShade) Then
        shadingDictionary.Add cellToShade, shadingValue
    Else
        shadingDictionary(cellToShade) = shadingValue
    End If
End Sub

Public Sub ApplyShading(ByVal ws As Worksheet, ByVal shadingDictionary As Object)
    Dim shadingKey As Variant
    Dim currentCell As Range

    For Each shadingKey In shadingDictionary.Keys
        ws.Range(shadingKey).Interior.Color = shadingDictionary(shadingKey)
    Next shadingKey
End Sub
