Private Sub Worksheet_Change(ByVal Target As Range)
    Dim changedCell As Range, validRange As Range

    Application.EnableEvents = False
    
    Set validRange = Union(Me.Range("C8:C32"), Me.Range("D8:D32"), _
                           Me.Range("E8:E32"), Me.Range("F8:F32"), _
                           Me.Range("G8:G32"), Me.Range("H8:H32"))
    
	' Allow the user to more easily enter a number or lowercase letter
    If Not Intersect(Target, validRange) Is Nothing Then
        For Each changedCell In Target
			' Automatically convert the number into a letter grade
            If IsNumeric(changedCell.Value) Then
                Select Case changedCell.Value
                    Case 1: changedCell.Value = "C"
                    Case 2: changedCell.Value = "B"
                    Case 3: changedCell.Value = "B+"
                    Case 4: changedCell.Value = "A"
                    Case 5: changedCell.Value = "A+"
                End Select
			' Automatically capitalize lowercase letters
            ElseIf VarType(changedCell.Value) = vbString Then
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