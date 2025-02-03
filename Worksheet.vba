Private Sub Worksheet_Change(ByVal targetCellsRange As Range)
    Dim changedCell As Range, englishNameRange As Range, koreanNameRange As Range
    Dim evalDateRange As Range, gradesRange As Range, commentRange As Range
    Dim cellValue As String
    
    On Error GoTo ErrorHandler
    Application.EnableEvents = False
    
    Set evalDateRange = Me.Range("C6")
    Set englishNameRange = Me.Range("B8:B32")
    Set koreanNameRange = Me.Range("C8:C32")
    Set gradesRange = Union(Me.Range("D8:D32"), Me.Range("E8:E32"), Me.Range("F8:F32"), _
                            Me.Range("G8:G32"), Me.Range("H8:H32"), Me.Range("I8:I32"))
    Set commentRange = Me.Range("J8:J32")
    
    For Each changedCell In targetCellsRange
        If Not IsEmpty(changedCell) Then changedCell.Value = Trim(changedCell.Value)
        
        If IsEmpty(changedCell) Or changedCell.Value = "" Then
            Select Case True
                Case Not Intersect(changedCell, englishNameRange) Is Nothing
                    SetDefaultCellShading changedCell, Me.Cells(7, 2).Value
                Case Not Intersect(changedCell, koreanNameRange) Is Nothing
                    SetDefaultCellShading changedCell, Me.Cells(7, 3).Value
                Case Not Intersect(changedCell, commentRange) Is Nothing
                    SetDefaultCellShading changedCell, Me.Cells(7, 10).Value
            End Select
        End If
        
        If changedCell.Value <> "" Then
            Select Case True
                Case Not Intersect(changedCell, evalDateRange) Is Nothing
                    ValdateEvalDateValue changedCell
                Case Not Intersect(changedCell, englishNameRange) Is Nothing
                    ValdateNameValue "English", changedCell
                Case Not Intersect(changedCell, koreanNameRange) Is Nothing
                    ValdateNameValue "Korean", changedCell
                Case Not Intersect(changedCell, gradesRange) Is Nothing
                    ValdateGradesValue changedCell
                    If changedCell.Value = "" Then changedCell.Select
                Case Not Intersect(changedCell, commentRange) Is Nothing
                    ValdateCommentValue changedCell
            End Select
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

Private Sub ValdateEvalDateValue(ByRef changedCell As Range)
    Dim cellValue As String, dateFormatStyle As String
    Dim userChoice As Integer
    
    cellValue = changedCell.Value
    
    Select Case Application.International(xlDateOrder)
       Case 0
           dateFormatStyle = "MM/DD/YYYY"
       Case 1
           dateFormatStyle = "DD/MM/YYYY"
       Case 2
           dateFormatStyle = "YYYY/MM/DD"
    End Select
    
    If IsDate(cellValue) Then
         changedCell.Value = Format(CDate(cellValue), "MMM. YYYY")
    Else
        msgToDisplay = "Please enter a valid date using " & dateFormatStyle & " formatting."
        userChoice = ThisWorkbook.DisplayMessage(msgToDisplay, vbOKOnly, "Invalid Date!", 200)
        changedCell.Select
        changedCell.Value = ""
    End If
End Sub

Private Sub ValdateNameValue(ByVal nameLanguage As String, ByRef changedCell As Range)
    Dim msgToDisplay As String, cellValue As String
    Dim userChoice As Integer
    
    cellValue = changedCell.Value
    
    Select Case nameLanguage
        Case "English"
            If Len(cellValue) > 30 Then
                If changedCell.Interior.Color <> RGB(255, 255, 0) Then UpdateCellShading changedCell, RGB(255, 255, 0)
                msgToDisplay = "The student's English name is longer than 30 characters and might not " & _
                               "fit on the report. Please verify how it looks after generating " & _
                               "the report and consider using a shorter version."
                userChoice = ThisWorkbook.DisplayMessage(msgToDisplay, vbOKOnly, "English Name Is Too Long!", 370)
                changedCell.Select
            Else
                If changedCell.Interior.Color <> xlNone Then UpdateCellShading changedCell, xlNone
            End If
        Case "Korean"
            If Len(cellValue) = 1 Or Len(cellValue) > 4 Then
                If changedCell.Interior.Color <> RGB(255, 0, 0) Then UpdateCellShading changedCell, RGB(255, 0, 0)
                msgToDisplay = "Please verify the you have written student's Korean name in Korean and spelled it correctly. An invalid length has been detected."
                userChoice = ThisWorkbook.DisplayMessage(msgToDisplay, vbOKOnly, "Name Error!", 350)
                changedCell.Select
            ElseIf Len(cellValue) = 2 Or Len(cellValue) = 4 Then
                If changedCell.Interior.Color <> RGB(255, 255, 0) Then UpdateCellShading changedCell, RGB(255, 255, 0)
                msgToDisplay = "Please verify the student's Korean name is correct. If you have typed it in English, please write it in Korean. " & _
                               "If you have written it in Korean, please check the spelling; a name of this length is uncommon."
                userChoice = ThisWorkbook.DisplayMessage(msgToDisplay, vbOKOnly, "Possible Name Error!", 380)
                changedCell.Select
            Else
                If changedCell.Interior.Color <> xlNone Then UpdateCellShading changedCell, xlNone
            End If
    End Select
End Sub

Private Sub ValdateGradesValue(ByRef changedCell As Range)
    Dim cellValue As String: cellValue = changedCell.Value
    
    If IsNumeric(cellValue) Then
        Select Case cellValue
            Case 1: changedCell.Value = "C"
            Case 2: changedCell.Value = "B"
            Case 3: changedCell.Value = "B+"
            Case 4: changedCell.Value = "A"
            Case 5: changedCell.Value = "A+"
            Case Else
                invalidScoreWarning
                changedCell.Value = ""
        End Select
    ElseIf VarType(cellValue) = vbString Then
        Select Case LCase(cellValue)
            Case "c": changedCell.Value = "C"
            Case "b": changedCell.Value = "B"
            Case "b+": changedCell.Value = "B+"
            Case "a": changedCell.Value = "A"
            Case "a+": changedCell.Value = "A+"
            Case Else
                If Len(cellValue) = 1 Then
                    invalidScoreWarning
                    changedCell.Value = ""
                ElseIf Len(cellValue) > 1 Then
                    TrimToLetterGrade changedCell
                End If
        End Select
    End If
End Sub

Private Sub TrimToLetterGrade(ByRef changedCell As Range)
    Dim firstCharacter As String, firstTwoCharacters As String, outsideCharacters As String
    
    firstTwoCharacters = UCase(Left(changedCell.Value, 2))
    firstCharacter = Left(firstTwoCharacters, 1)
    outsideCharacters = firstCharacter & Right(changedCell.Value, 1)
    
    Select Case True
        Case (firstTwoCharacters = "A+"), (firstTwoCharacters = "B+")
            changedCell.Value = firstTwoCharacters
        Case (outsideCharacters = "A+"), (outsideCharacters = "B+")
            changedCell.Value = outsideCharacters
        Case (firstCharacter = "A"), (firstCharacter = "B"), (firstCharacter = "C")
            changedCell.Value = firstCharacter
        Case Else
            invalidScoreWarning
            changedCell.Value = ""
    End Select
End Sub

Private Sub invalidScoreWarning()
    Const MSG_TO_DISPLAY As String = "An invalid score value has been entered. Please try entering the score again."
    Dim userChoice As Integer

    userChoice = ThisWorkbook.DisplayMessage(MSG_TO_DISPLAY, vbOKOnly, "Invalid Value!", 250)
End Sub

Private Sub ValdateCommentValue(ByRef changedCell As Range)
    Dim msgToDisplay As String, userChoice As Integer
    Dim cellValue As String
    
    cellValue = changedCell.Value
    
    Select Case True
        Case Len(cellValue) = 0
            If changedCell.Interior.Color <> RGB(242, 242, 242) Then UpdateCellShading changedCell, RGB(242, 242, 242)
        Case Len(cellValue) < 80
            If changedCell.Interior.Color <> RGB(255, 255, 0) Then UpdateCellShading changedCell, RGB(255, 255, 0)
            msgToDisplay = "The comment you have typed is very short. Please check that you " & _
                           "have followed the ""Positive - Negative - Positive"" format."
            userChoice = ThisWorkbook.DisplayMessage(msgToDisplay, vbOKOnly, "Short Comment!", 250)
            changedCell.Select
        Case Len(cellValue) > 315
            If changedCell.Interior.Color <> RGB(255, 0, 0) Then UpdateCellShading changedCell, RGB(255, 0, 0)
            msgToDisplay = "The comment you have typed is too long. Please shorten it by " & _
                           Len(cellValue) - 315 & " characters (or more) to ensure it properly " & _
                           "fits in the report's comment box."
            userChoice = ThisWorkbook.DisplayMessage(msgToDisplay, vbOKOnly, "Long Comment!", 250)
            changedCell.Select
        Case Else
            If changedCell.Interior.Color <> RGB(242, 242, 242) Then UpdateCellShading changedCell, RGB(242, 242, 242)
    End Select
End Sub

Private Sub SetDefaultCellShading(ByRef changedCell As Range, ByVal columnName As String)
    Dim shadingColour As Long
    
    Select Case columnName
        Case "English Name", "Korean Name"
            shadingColour = xlNone
        Case "Comments"
            shadingColour = RGB(242, 242, 242)
    End Select
    
    If changedCell.Interior.Color <> shadingColour Then UpdateCellShading changedCell, shadingColour
End Sub

Private Sub UpdateCellShading(ByRef changedCell As Range, ByVal shadingColour As Long)
    Me.Unprotect
    changedCell.Interior.Color = shadingColour
    Me.Protect
    Me.EnableSelection = xlUnlockedCells
End Sub
