Private Sub Worksheet_Change(ByVal targetCellsRange As Range)
    Dim changedCell As Range, englishNameRange As Range, koreanNameRange As Range
    Dim gradesRange As Range, commentRange As Range
    Dim cellValue As String
    
    On Error GoTo ErrorHandler
    Application.EnableEvents = False
    
    Set englishNameRange = Me.Range("B8:B32")
    Set koreanNameRange = Me.Range("C8:C32")
    Set gradesRange = Union(Me.Range("D8:D32"), Me.Range("E8:E32"), Me.Range("F8:F32"), _
                            Me.Range("G8:G32"), Me.Range("H8:H32"), Me.Range("I8:I32"))
    Set commentRange = Me.Range("J8:J32")
    
    For Each changedCell In targetCellsRange
        If changedCell.Value <> "" Then
           cellValue = Trim(changedCell.Value)
            
            Select Case True
                Case Not Intersect(changedCell, englishNameRange) Is Nothing
                    ValdateNameValue "English", changedCell, cellValue
                Case Not Intersect(changedCell, koreanNameRange) Is Nothing
                    ValdateNameValue "Korean", changedCell, cellValue
                Case Not Intersect(changedCell, gradesRange) Is Nothing
                    ValdateGradesValue cellValue
                Case Not Intersect(changedCell, commentRange) Is Nothing
                    ValdateCommentValue changedCell, cellValue
            End Select
            
            If cellValue = "" Then changedCell.Select
            changedCell.Value = cellValue
        End If
    Next changedCell

CleanUp:
    Application.EnableEvents = True
    Exit Sub
ErrorHandler:
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Error message to be added."
    #End If
    Resume CleanUp
End Sub

Private Sub ValdateNameValue(ByVal nameLanguage As String, ByRef changedCell As Range, ByRef cellValue As String)
    Dim msgToDisplay As String, userChoice As Integer
    
    Select Case nameLanguage
        Case "English"
            If Len(cellValue) > 30 Then
                changedCell.Interior.Color = RGB(255, 255, 0)
                msgToDisplay = "The student's English name is longer than 30 characters and might not " & _
                               "fit on the report. Please verify how it looks after generating " & _
                               "the report and consider using a shorter version."
                userChoice = ThisWorkbook.DisplayMessage(msgToDisplay, vbOKOnly, "English Name Is Too Long!", 250)
                changedCell.Select
            Else
                changedCell.Interior.ColorIndex = xlNone
            End If
        Case "Korean"
            ' If possible, add a check to see English letters, numbers, or punctuation is present.
            If Len(cellValue) = 1 Or Len(cellValue) > 4 Then
                changedCell.Interior.Color = RGB(255, 0, 0)
                msgToDisplay = "Please verify the student's Korean name is correct. An invalid length has been detected."
                userChoice = ThisWorkbook.DisplayMessage(msgToDisplay, vbOKOnly, "Name Error!", 250)
                changedCell.Select
            ElseIf Len(cellValue) = 2 Or Len(cellValue) = 4 Then
                changedCell.Interior.Color = RGB(255, 255, 0)
                msgToDisplay = "Please verify the student's Korean name is correct. While some students " & _
                               "have a name of this length, it is very uncommon."
                userChoice = ThisWorkbook.DisplayMessage(msgToDisplay, vbOKOnly, "Possible Name Error!", 250)
                changedCell.Select
            Else
                changedCell.Interior.ColorIndex = xlNone
            End If
    End Select
End Sub

Private Sub ValdateGradesValue(ByRef cellValue As String)
    If IsNumeric(cellValue) Then
        Select Case cellValue
            Case 1: cellValue = "C"
            Case 2: cellValue = "B"
            Case 3: cellValue = "B+"
            Case 4: cellValue = "A"
            Case 5: cellValue = "A+"
            Case Else: invalidScoreValue cellValue
        End Select
    ElseIf VarType(cellValue) = vbString Then
        Select Case LCase(cellValue)
            Case "c": cellValue = "C"
            Case "b": cellValue = "B"
            Case "b+": cellValue = "B+"
            Case "a": cellValue = "A"
            Case "a+": cellValue = "A+"
            Case Else
                If Len(cellValue) = 1 Then
                    invalidScoreValue cellValue
                ElseIf Len(cellValue) > 1 Then
                    TrimToLetterGrade cellValue
                End If
        End Select
    End If
End Sub

Private Sub TrimToLetterGrade(ByRef cellValue As String)
    Dim firstCharacter As String, firstTwoCharacters As String, outsideCharacters As String
    
    firstTwoCharacters = UCase(Left(cellValue, 2))
    firstCharacter = Left(firstTwoCharacters, 1)
    outsideCharacters = firstCharacter & Right(cellValue, 1)
    
    Select Case True
        Case (firstTwoCharacters = "A+"), (firstTwoCharacters = "B+")
            cellValue = firstTwoCharacters
        Case (outsideCharacters = "A+"), (outsideCharacters = "B+")
            cellValue = outsideCharacters
        Case (firstCharacter = "A"), (firstCharacter = "B"), (firstCharacter = "C")
            cellValue = firstCharacter
        Case Else
            invalidScoreValue cellValue
    End Select
End Sub

Private Sub invalidScoreValue(ByRef cellValue As String)
    Const MSG_TO_DISPLAY As String = "An invalid score value has been entered. Please try entering the score again."
    Dim userChoice As Integer

    userChoice = ThisWorkbook.DisplayMessage(MSG_TO_DISPLAY, vbOKOnly, "Invalid Value!", 250)
    cellValue = ""
End Sub

Private Sub ValdateCommentValue(ByRef changedCell As Range, ByRef cellValue As String)
    Dim msgToDisplay As String, userChoice As Integer
    
    Select Case True
        Case Len(cellValue) < 80
            changedCell.Interior.Color = RGB(255, 255, 0)
            msgToDisplay = "The comment you have typed is very short. Please check that you " & _
                           "have followed the ""Positive - Negative - Positive"" format."
            userChoice = ThisWorkbook.DisplayMessage(msgToDisplay, vbOKOnly, "Short Comment!", 250)
            changedCell.Select
        Case Len(cellValue) > 315
            changedCell.Interior.Color = RGB(255, 0, 0)
            msgToDisplay = "The comment you have typed is too long. Please shorten it by " & _
                           Len(cellValue) - 315 & " characters (or more) to ensure it fits " & _
                           "properly in the reports comment box."
            userChoice = ThisWorkbook.DisplayMessage(msgToDisplay, vbOKOnly, "Long Comment!", 250)
            changedCell.Select
        Case Else
            changedCell.Interior.Color = RGB(242, 242, 242)
    End Select
End Sub
