Option Explicit

#Const PRINT_DEBUG_MESSAGES = True
#If Mac Then
    Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
    Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
#End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Record Updates
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateClassRecords(ByVal ws As Worksheet, ByVal targetCellsRange As Range)
    Dim currentCell As Range
    Dim cellToShade As Range
    Dim cellToQuery As Range
    Dim validationListRange As Range
    Dim winnersListRange As Range
    Dim englishNamesRange As Range
    Dim enteredValue() As String
    Dim validatedValue() As String
    Dim fieldType() As String
    Dim nameToFind As String
    Dim shadingValue As Long
    Dim totalChangedCells As Long
    Dim rankShading As Long
    Dim i As Long
    Dim currentRow As Long
    Dim studentNameUpdate As Boolean
    Dim winningName As Boolean
    Dim shadingUpdates As New Dictionary
    
    ' Step 1: Prepare dictionary and arrays
    ToggleApplicationFeatures False
    ToggleSheetProtection ws, False
    On Error GoTo ErrorHandler
    
    ' Set shadingUpdates = CreateObject("Scripting.Dictionary")
    With ws
        Set validationListRange = .Range(RANGE_VALIDATION_LIST)
        Set winnersListRange = .Range(RANGE_WINNERS)
        Set englishNamesRange = .Range(RANGE_ENGLISH_NAME)
    End With

    studentNameUpdate = False
    totalChangedCells = targetCellsRange.Cells.Count
    
    ReDim enteredValue(1 To totalChangedCells)
    ReDim validatedValue(1 To totalChangedCells)
    ReDim fieldType(1 To totalChangedCells)
    
    ' Step 2: Validate input and determine fieldType
    For i = 1 To totalChangedCells
        Set currentCell = targetCellsRange.Cells.Item(i)
        enteredValue(i) = Trim$(CStr(currentCell.Value))
        validatedValue(i) = enteredValue(i)
        fieldType(i) = GetCellType(currentCell)
        studentNameUpdate = False

        Select Case fieldType(i)
            Case "Native Teacher", "Korean Teacher"
                validatedValue(i) = FormatName(validatedValue(i))
            Case "English Name", "Korean Name"
                validatedValue(i) = FormatName(validatedValue(i))
                studentNameUpdate = True
            Case "Eval Date"
                validatedValue(i) = FormatEvalDate(validatedValue(i))
            Case "Grade"
                validatedValue(i) = FormatGrade(validatedValue(i))
            Case "Comment"
                validatedValue(i) = FormatComment(validatedValue(i))
        End Select

        ' Step 2a: Write new value if input has been updated
        If enteredValue(i) <> validatedValue(i) Then currentCell.Value = validatedValue(i)
        
        ' Step 2b: Update winners list options if a student's name has been updated
        If studentNameUpdate Then UpdateHiddenNameValidationList ws, currentCell
    Next i
    
    ' Step 3: Process shading for "English Name", "Korean Name", and "Comment" cells
    For i = 1 To totalChangedCells
        Set currentCell = targetCellsRange.Cells.Item(i)
        winningName = False
        shadingValue = 0

        ' Step 3a: Determine if current name has been selected as a winner
        If fieldType(i) = "English Name" Or fieldType(i) = "Korean Name" Then
            winningName = CheckIfAWinningName(ws, fieldType(i), currentCell)
        End If

        Select Case True
            ' Step 3b: Handle updates to the winners list
            Case Not Intersect(currentCell, winnersListRange) Is Nothing
                ' Step 3b-i: Load shading value for the student's rank
                shadingValue = GetWinnerShadingValue(currentCell.Address)

                If enteredValue(i) = vbNullString Then
                    ' Step 3b-ii: Set proper shading if name has been removed from the winners list
                    DetermineNameCellShading ws, englishNamesRange, shadingValue, shadingUpdates
                Else
                    ' Step 3b-iii: Queue shading for winner students
                    SetShadingForWinnerName validationListRange, enteredValue(i), shadingUpdates, shadingValue
                    
                    ' Step 3b-iv: Remove duplicates in L2:L4
                    RemoveDuplicateWinners ws, winnersListRange, currentCell.Address, shadingValue, shadingUpdates, enteredValue(i)
                
                    ' Step 3b-v: Queue shading for non-winning students
                    DetermineNameCellShading ws, englishNamesRange, shadingValue, shadingUpdates
                End If
            ' Step 3c: Handle updates to non-winning students
            Case Not winningName
                ' Step 3c-i: Queue default shading for empty/blank cells
                If IsEmpty(enteredValue(i)) Or enteredValue(i) = vbNullString Then
                    Select Case fieldType(i)
                        Case "English Name", "Korean Name"
                            AddToShadingDictionary shadingUpdates, currentCell.Address, xlNone
                        Case "Comment"
                            AddToShadingDictionary shadingUpdates, currentCell.Address, RGB(242, 242, 242)
                    End Select
                ' Step 3c-ii: Queue shading for updated cells
                Else
                    Select Case fieldType(i)
                        Case "English Name"
                            AddToShadingDictionary shadingUpdates, currentCell.Address, GetEnglishNameShading(validatedValue(i))
                        Case "Korean Name"
                            AddToShadingDictionary shadingUpdates, currentCell.Address, GetKoreanNameShading(validatedValue(i))
                        Case "Comment"
                            AddToShadingDictionary shadingUpdates, currentCell.Address, GetCommentShading(validatedValue(i))
                    End Select
                End If
        End Select
    Next i
    
    ' Step 4: Apply shading queue
    ApplyShading ws, shadingUpdates
    
' Step 5: Clear and reset settings.
CleanUp:
    ToggleApplicationFeatures True
    ToggleSheetProtection ws, True
    Set currentCell = Nothing
    Set shadingUpdates = Nothing
    Exit Sub
ErrorHandler:
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Error in Worksheet_Change: " & Err.Description & " (Error " & Err.Number & ")"
    #End If
    Resume CleanUp
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Cell and Data Updating
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckIfAWinningName(ByVal ws As Worksheet, ByVal cellType As String, ByVal currentCell As Range) As Boolean
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

            .Range("O" & i).Value = mergedName
        Next i
    End With
End Sub

Private Sub UpdateHiddenNameValidationList(ByVal ws As Worksheet, ByVal currentCell As Range)
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
        
        currentName = .Range("O" & currentRow).Value
        If currentName <> mergedName Then
            .Range("O" & currentRow).Value = mergedName
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
            If Not IsEmpty(.Range("O" & i)) Then
                validationList = validationList & .Range("O" & i).Value & ","
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

Private Function FormatName(ByVal nameValue As String) As String
    If IsValueEnglish(nameValue) Then
        FormatName = StrConv(nameValue, vbProperCase)
    Else
        FormatName = nameValue
    End If
End Function

Private Function IsValueEnglish(ByVal inputText As String) As Boolean
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

Private Function FormatComment(ByVal commentValue As String) As String
    FormatComment = UCase$(Left$(commentValue, 1)) & Mid$(commentValue, 2)
End Function

Private Function FormatEvalDate(ByVal dateValue As String) As String
    Dim msgTitle As String
    Dim msgToDisplay As String
    Dim msgDialogType As Long
    Dim msgDialogWidth As Long
    
    If IsDate(dateValue) Then
        ' Make this format match the user's locale
        FormatEvalDate = Format$(CDate(dateValue), "DD MMM. YYYY")
    Else
        If dateValue <> vbNullString Then
            msgTitle = "Date: Invalid Format"
            msgToDisplay = "Please enter a valid date."
            msgDialogType = vbOKOnly + vbExclamation
            msgDialogWidth = 200
            Call DisplayMessage(msgToDisplay, msgDialogType, msgTitle, msgDialogWidth)
        End If
        FormatEvalDate = vbNullString
    End If
End Function

Private Function FormatGrade(ByVal gradeValue As String) As String
    Dim msgTitle As String
    Dim msgToDisplay As String
    Dim msgDialogType As Long
    Dim msgDialogWidth As Long
    Dim processedGrade As String
    
    processedGrade = UCase$(Trim$(gradeValue))
    
    If processedGrade <> vbNullString Then
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
                 
                 If processedGrade = vbNullString Then
                    msgTitle = "Grade: Invalid Score"
                    msgToDisplay = "An invalid score value has been entered. Please enter A+, A, B+, B, C, or a number between 1 and 5."
                    msgDialogType = vbOKOnly + vbExclamation
                    msgDialogWidth = 250
                    Call DisplayMessage(msgToDisplay, msgDialogType, msgTitle, msgDialogWidth)
                 End If
                 
                 FormatGrade = processedGrade
        End Select
    End If
End Function

Private Function TrimToFinalLetterGrade(ByVal inputText As String) As String
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
