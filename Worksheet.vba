Private Sub Worksheet_Change(ByVal targetCellsRange As Range)
    Dim levelRange As Range, classDaysRange As Range, classTimeRange As Range, evalDateRange As Range
    Dim englishNameRange As Range, gradesRange As Range, commentRange As Range
    Dim changedCell As Range
    
    Application.EnableEvents = False
    
    targetCellsRange.Value = Trim(targetCellsRange.Value)
    If targetCellsRange.Value = "" Then Exit Sub
    
    Set englishNameRange = Me.Range("B8:B32")
    Set gradesRange = Union(Me.Range("D8:D32"), Me.Range("E8:E32"), Me.Range("F8:F32"), _
                            Me.Range("G8:G32"), Me.Range("H8:H32"), Me.Range("I8:I32"))
    Set commentRange = Me.Range("J8:J32")
    
    If Not Intersect(targetCellsRange, englishNameRange) Is Nothing Then
        ValdateEnglishNameValue targetCellsRange
    ElseIf Not Intersect(targetCellsRange, gradesRange) Is Nothing Then
        ValdateGradesValue targetCellsRange
    ElseIf Not Intersect(targetCellsRange, commentRange) Is Nothing Then
        ValdateCommentValue targetCellsRange
    End If

    Application.EnableEvents = True
End Sub

Private Sub ValdateEnglishNameValue(ByVal targetCell As Range)
     ' Add a character limit check
     ' Add a prompt to ask to trim to character limit
End Sub

Private Sub ValdateGradesValue(ByVal targetCell As Range)
    Dim changedCell As Range

    For Each changedCell In targetCell
        If IsNumeric(changedCell.Value) Then
            Select Case changedCell.Value
                Case 1: changedCell.Value = "C"
                Case 2: changedCell.Value = "B"
                Case 3: changedCell.Value = "B+"
                Case 4: changedCell.Value = "A"
                Case 5: changedCell.Value = "A+"
                Case Else: invalidScoreValue changedCell
            End Select
        ElseIf VarType(changedCell.Value) = vbString Then
            Select Case LCase(changedCell.Value)
                Case "c": changedCell.Value = "C"
                Case "b": changedCell.Value = "B"
                Case "b+": changedCell.Value = "B+"
                Case "a": changedCell.Value = "A"
                Case "a+": changedCell.Value = "A+"
                Case Else
                    If Len(changedCell.Value) = 1 Then invalidScoreValue changedCell
                    If Len(changedCell.Value) > 1 Then TrimToLetterGrade changedCell
            End Select
        End If
    Next changedCell
End Sub

Private Sub TrimToLetterGrade(ByVal changedCell As Range)
    Dim firstCharacter As String, startingCharacters As String, outsideCharacters As String
    
    startingCharacters = UCase(Left(changedCell.Value, 2))
    firstCharacter = Left(startingCharacters, 1)
    outsideCharacters = firstCharacter & Right(changedCell.Value, 1)
    
    If startingCharacters = "A+" Or startingCharacters = "B+" Then
        changedCell.Value = startingCharacters
    ElseIf outsideCharacters = "A+" Or outsideCharacters = "B+" Then
        changedCell.Value = outsideCharacters
    ElseIf firstCharacter = "A" Or firstCharacter = "B" Or firstCharacter = "C" Then
        changedCell.Value = firstCharacter
    Else
        invalidScoreValue changedCell
    End If
End Sub

Private Sub invalidScoreValue(ByVal changedCell As Range, Optional ByVal wrongNumber As Integer)
    Const MSG_TO_DISPLAY As String = "An invalid score value has been entered. Please try entering the score again."
    Dim userChoice As Integer

    userChoice = ThisWorkbook.DisplayMessage(MSG_TO_DISPLAY, vbRetryCancel, "Invalid Value!", 250)
    
    #If Mac Then
        changedCell.Select
        changedCell.Value = ""
    #End If
End Sub

Private Sub ValdateCommentValue(ByVal targetCell As Range)
    ' Add a character limit check
    ' Can be grabbed from the check when prepping the reports
End Sub
