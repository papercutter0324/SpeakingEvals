Option Explicit

#Const PRINT_DEBUG_MESSAGES = True
#If Mac Then
    Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
    Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
#End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Cell and Data Updating
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckIfAWinningName(ByVal ws As Worksheet, ByVal cellType As String, ByVal currentCell As Range) As Boolean
    Dim englishName As String
    Dim koreanName As String
    Dim mergedName As String
    Dim matchFound As Boolean
    Dim i As Long
    
    Select Case cellType
        Case "English Name"
            i = 1
        Case "Korean Name"
            i = -1
        Case Else
            CheckIfAWinningName = False
            Exit Function
    End Select

    If IsEmpty(currentCell) Or IsEmpty(currentCell.Offset(0, i)) Then
        CheckIfAWinningName = False
        Exit Function
    End If
    
    With currentCell
        Select Case cellType
            Case "English Name"
                englishName = .Value
                koreanName = .Offset(0, i).Value
            Case "Korean Name"
                englishName = .Offset(0, i).Value
                koreanName = .Value
        End Select
    End With
            
    mergedName = englishName & "(" & koreanName & ")"
    matchFound = False
    
    For i = 2 To 4
        If ws.Range("L" & i).Value = mergedName Then
            matchFound = True
            Exit For
        End If
    Next i
    
    CheckIfAWinningName = matchFound
End Function

Public Sub GenerateCompleteHiddenNameValidationList(ByVal ws As Worksheet)
    Dim mergedName As String
    Dim i As Long

    With ws
        For i = 8 To 32
            If Not IsEmpty(.Range("B" & i)) And Not IsEmpty(.Range("C" & i)) Then
                mergedName = .Range("B" & i).Value & "(" & .Range("C" & i).Value & ")"
            Else
                mergedName = vbNullString
            End If

            .Range("BB" & i).Value = mergedName
        Next i
    End With
End Sub

Public Sub UpdateHiddenNameValidationList(ByVal ws As Worksheet, ByVal currentCell As Range)
    Dim offsetValue As Long
    Dim currentRow As Long
    Dim currentName As String
    Dim mergedName As String

    offsetValue = IIf(currentCell.Column = 2, 1, -1)
    currentRow = currentCell.Row
    
    With ws
        If Not IsEmpty(currentCell) And Not IsEmpty(currentCell.Offset(0, offsetValue)) Then
            mergedName = .Range("B" & currentRow).Value & "(" & .Range("C" & currentRow).Value & ")"
        Else
            mergedName = vbNullString
        End If
        
        currentName = .Range("BB" & currentRow).Value
        If currentName <> mergedName Then
            .Range("BB" & currentRow).Value = mergedName
            PopulateWinnersListValidationValues ws
        End If
    End With
End Sub

Public Sub PopulateWinnersListValidationValues(ByVal ws As Worksheet)
    Dim dataValidation As Object
    Dim validationList As String
    Dim i As Long

    With ws
        For i = 8 To 32
            If Not IsEmpty(.Range("BB" & i)) Then
                validationList = validationList & .Range("BB" & i).Value & ","
            End If
        Next i

        If validationList = vbNullString Then
            validationList = "Incomplete List"
        ElseIf Right$(validationList, 1) = "," Then
            validationList = Left$(validationList, Len(validationList) - 1)
        End If
        
        ' Clear existing data validation
        On Error Resume Next
        .Range("L2:L4").Validation.Delete
        On Error GoTo 0

        ' Apply data validation to winners range
        With .Range("L2:L4").Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=validationList
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    End With
End Sub

Public Function FormatName(ByVal nameValue As String) As String
    If IsValueEnglish(nameValue) Then
        FormatName = StrConv(nameValue, vbProperCase)
    Else
        FormatName = nameValue
    End If
End Function

Public Function IsValueEnglish(ByVal inputText As String) As Boolean
    Dim charCode As Long
    Dim i As Long
    
    For i = 1 To Len(inputText)
        charCode = AscW(Mid$(inputText, i, 1)) ' Use AscW for Unicode characters

        ' Option 1: Broad check characters if are not standard ASCII characters
        ' If charCode > 127 Then
        '     IsValueEnglish = False
        '     Exit Function
        ' End If
        
        ' Option 2: Specifically check for Hangul Syllables (Unicode range: AC00 to D7AF (Hex) OR 44032 to 55215 (Decimal))
        If charCode >= &HAC00 And charCode <= &HD7AF Then
            IsValueEnglish = False
            Exit Function
        End If
    Next i
    
    IsValueEnglish = True
End Function

Public Function FormatComment(ByVal commentValue As String) As String
    FormatComment = UCase$(Left$(commentValue, 1)) & Mid$(commentValue, 2)
End Function

Public Function FormatEvalDate(ByVal dateValue As String) As String
    If IsDate(dateValue) Then
        FormatEvalDate = Format$(CDate(dateValue), "MMM. YYYY")
    Else
        DisplayWarning "Date: Invalid Format"
        FormatEvalDate = vbNullString
    End If
End Function

Public Function FormatGrade(ByVal gradeValue As String) As String
    Dim processedGrade As String
    
    processedGrade = UCase$(Trim$(gradeValue))

    Select Case processedGrade
        Case "A+", "A", "B+", "B", "C"
            FormatGrade = processedGrade
        Case "1": FormatGrade = "C"
        Case "2": FormatGrade = "B"
        Case "3": FormatGrade = "B+"
        Case "4": FormatGrade = "A"
        Case "5": FormatGrade = "A+"
        Case Else
            ' If no direct match, attempt to determine intended grade
             processedGrade = TrimToFinalLetterGrade(processedGrade)
             If processedGrade <> vbNullString Then
                 FormatGrade = processedGrade
             Else
                DisplayWarning "Grade: Invalid Score"
                FormatGrade = vbNullString
             End If
    End Select
End Function

Public Function TrimToFinalLetterGrade(ByVal inputText As String) As String
    Dim firstChar As String
    Dim firstTwo As String
    Dim outsideTwo As String
    Dim cleanedInput As String
    Dim trimSuccessful As Boolean

    TrimToFinalLetterGrade = vbNullString

    If Len(inputText) = 0 Then Exit Function

    firstChar = Left$(inputText, 1)
    
    If Len(inputText) >= 2 Then
        firstTwo = Left$(inputText, 2)
        outsideTwo = firstChar & Right$(inputText, 1)
    Else
        firstTwo = vbNullString
        outsideTwo = vbNullString
    End If
    
    ' Check specific valid formats
    Select Case True
        Case (firstTwo = "A+"), (firstTwo = "B+")
            TrimToFinalLetterGrade = firstTwo
            trimSuccessful = True
        Case (outsideTwo = "A+"), (outsideTwo = "B+")
            TrimToFinalLetterGrade = outsideTwo
            trimSuccessful = True
        Case (firstChar = "A"), (firstChar = "B"), (firstChar = "C")
            TrimToFinalLetterGrade = firstChar
            trimSuccessful = True
    End Select

    ' Add simple cleanup for common errors like "A +" -> "A+"
    If Not trimSuccessful = True And Len(inputText) > 1 Then
         cleanedInput = Replace(inputText, " ", vbNullString) ' Remove internal spaces
         Select Case cleanedInput
             Case "A+", "B+"
                TrimToFinalLetterGrade = cleanedInput
                Exit Function
         End Select
    End If
End Function
