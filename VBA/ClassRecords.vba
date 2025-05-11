Option Explicit

#Const PRINT_DEBUG_MESSAGES = True
#If Mac Then
    Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
    Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
#End If

Private Sub Worksheet_Change(ByVal targetCellsRange As Range)
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
    Dim shadingUpdates As Object
    
    ' Step 1: Prepare dictionary and arrays
    Me.Unprotect
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    On Error GoTo ErrorHandler
    
    Set shadingUpdates = CreateObject("Scripting.Dictionary")
    Set validationListRange = Me.Range(RANGE_VALIDATION_LIST)
    Set winnersListRange = Me.Range(RANGE_WINNERS)
    Set englishNamesRange = Me.Range(RANGE_ENGLISH_NAME)

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
        If studentNameUpdate Then UpdateHiddenNameValidationList Me, currentCell
    Next i
    
    ' Step 3: Process shading for "English Name", "Korean Name", and "Comment" cells
    For i = 1 To totalChangedCells
        Set currentCell = targetCellsRange.Cells.Item(i)
        winningName = False
        shadingValue = 0

        ' Step 3a: Determine if current name has been selected as a winner
        If fieldType(i) = "English Name" Or fieldType(i) = "Korean Name" Then
            winningName = CheckIfAWinningName(Me, fieldType(i), currentCell)
        End If

        Select Case True
            ' Step 3b: Handle updates to the winners list
            Case Not Intersect(currentCell, winnersListRange) Is Nothing
                ' Step 3b-i: Load shading value for the student's rank
                shadingValue = GetWinnerShadingValue(currentCell.Address)

                If enteredValue(i) = vbNullString Then
                    ' Step 3b-ii: Set proper shading if name has been removed from the winners list
                    DetermineNameCellShading Me, englishNamesRange, shadingValue, shadingUpdates
                Else
                    ' Step 3b-iii: Queue shading for winner students
                    SetShadingForWinnerName validationListRange, enteredValue(i), shadingUpdates, shadingValue
                    
                    ' Step 3b-iv: Remove duplicates in L2:L4
                    RemoveDuplicateWinners Me, winnersListRange, currentCell.Address, shadingValue, shadingUpdates, enteredValue(i)
                
                    ' Step 3b-v: Queue shading for non-winning students
                    DetermineNameCellShading Me, englishNamesRange, shadingValue, shadingUpdates
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
    ApplyShading Me, shadingUpdates
    
' Step 5: Clear and reset settings.
CleanUp:
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    With Me
        .Protect
        .EnableSelection = xlUnlockedCells
    End With
    Set currentCell = Nothing
    Set shadingUpdates = Nothing
    Exit Sub
ErrorHandler:
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Error in Worksheet_Change: " & Err.Description & " (Error " & Err.Number & ")"
    #End If
    Resume CleanUp
End Sub
