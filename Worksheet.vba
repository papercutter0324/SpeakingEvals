Private Sub Worksheet_Change(ByVal Target As Range)
    Dim changedCell As Range, validRange As Range

    Application.EnableEvents = False
    
    ' Monitor the six letter grade columns
    Set validRange = Union(Me.Range("D8:D32"), Me.Range("E8:E32"), _
                           Me.Range("F8:F32"), Me.Range("G8:G32"), _
                           Me.Range("H8:H32"), Me.Range("I8:I32"))
    
    If Not Intersect(Target, validRange) Is Nothing Then
        For Each changedCell In Target
            If IsNumeric(changedCell.Value) Then
                ' Convert numbers into letter grades
		Select Case changedCell.Value
                    Case 1: changedCell.Value = "C"
                    Case 2: changedCell.Value = "B"
                    Case 3: changedCell.Value = "B+"
                    Case 4: changedCell.Value = "A"
                    Case 5: changedCell.Value = "A+"
                End Select
            ElseIf VarType(changedCell.Value) = vbString Then
		' Ensure letter grades are capitalized
                Select Case LCase(changedCell.Value)
                    Case "c": changedCell.Value = "C"
                    Case "b": changedCell.Value = "B"
                    Case "b+": changedCell.Value = "B+"
                    Case "a": changedCell.Value = "A"
                    Case "a+": changedCell.Value = "A+"
                End Select
            End If
        Next changedCell
    End If

    Application.EnableEvents = True
End Sub
