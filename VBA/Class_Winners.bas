Option Explicit

#Const Windows = (Mac = 0)

Public Sub AutoSelectClassWinners(ByRef ws As Worksheet)
    Dim nameList()  As String
    Dim scoreList() As Double
    Dim lastRow     As Long

    ' Determine the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Resize arrays to hold names and scores
    ReDim nameList(1 To lastRow - g_STUDENT_INDEX_OFFSET)
    GenerateNameList ws, nameList(), g_STUDENT_INDEX_OFFSET + 1, lastRow

    If nameList(LBound(nameList)) = vbNullString Then
        ' Spit out an error
        Exit Sub
    End If

    ReDim scoreList(1 To UBound(nameList))
    GenerateScoreList ws, scoreList()

    BubbleSortTwoLists nameList(), scoreList()

    PopulateWinnerSelections ws, nameList()
    SetWinnersListValidation ws, nameList()

    ToggleSheetProtection ws, False
    SetDefaultShading ws
    ToggleSheetProtection ws, True
End Sub

Public Sub UpdateWinnersLists(ByRef ws As Worksheet, Optional ByVal autoSelectWinners As Boolean = False)
    Dim nameList()  As String
    Dim scoreList() As Double
    Dim lastRow     As Long

    If autoSelectWinners Then
        ToggleSheetProtection ws, False
    End If

    ' Determine the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Resize arrays to hold names and scores
    ReDim nameList(1 To lastRow - g_STUDENT_INDEX_OFFSET)
    GenerateNameList ws, nameList(), g_STUDENT_INDEX_OFFSET + 1, lastRow

    If nameList(LBound(nameList)) = vbNullString Then
        ' Spit out an error
        Exit Sub
    End If

    ReDim scoreList(1 To UBound(nameList))
    GenerateScoreList ws, scoreList()

    BubbleSortTwoLists nameList(), scoreList()

    If autoSelectWinners Then
        PopulateWinnerSelections ws, nameList()
    End If
    
    SetWinnersListValidation ws, nameList()

    If autoSelectWinners Then
        SetDefaultShading ws
        ToggleSheetProtection ws, True
    End If
End Sub

Private Sub GenerateNameList(ByRef ws As Worksheet, ByRef nameList() As String, ByVal startRow As Long, ByVal endRow As Long)
    Const ENG_NAME_COL As Long = 2
    Const KOR_NAME_COL As Long = 3
    
    Dim rowIndex      As Long
    Dim nameListIndex As Long
    Dim engName       As String
    Dim korName       As String

    With ws
        For rowIndex = startRow To endRow
            If Not IsEmpty(.Cells(rowIndex, ENG_NAME_COL)) And Not IsEmpty(.Cells(rowIndex, KOR_NAME_COL)) Then
                nameListIndex = nameListIndex + 1
                engName = Trim$(.Cells(rowIndex, ENG_NAME_COL).Value)
                korName = Trim$(.Cells(rowIndex, KOR_NAME_COL).Value)
                
                nameList(nameListIndex) = engName & " (" & korName & ")"
            End If
        Next rowIndex
    End With
    
    If nameListIndex > 0 Then
        ReDim Preserve nameList(1 To nameListIndex)
    End If
End Sub

Public Sub SetWinnersListValidation(ByRef ws As Worksheet, ByRef nameList() As String)
    Dim validationValues As ValidationSettings
    Dim validationList  As String
    Dim i               As Long

    For i = LBound(nameList) To UBound(nameList)
        validationList = validationList & nameList(i) & ","
    Next i

    If Len(validationList) > 0 Then
        validationList = Left$(validationList, Len(validationList) - 1)
    End If

    With validationValues
        .TypeOfValidation = xlValidateList
        .AlertStyle = xlValidAlertStop
        .Operator = xlBetween
        .InputTitle = vbNullString
        .InputMessage = vbNullString
        .Formula1 = validationList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

    ApplyValidationValues ws.Range(g_WINNER_NAMES), validationValues
End Sub

Private Sub GenerateScoreList(ByRef ws As Worksheet, ByRef scoreList() As Double)
    Dim studentIndex        As Long
    Dim gradeCategoryIndex  As Long
    Dim allStudentsScores   As Variant

    allStudentsScores = ws.Range(g_STUDENT_GRADES).Value

    For studentIndex = LBound(scoreList) To UBound(scoreList)
        scoreList(studentIndex) = 0 ' Reset score in case of score changes

        For gradeCategoryIndex = 1 To 6
            Select Case allStudentsScores(studentIndex, gradeCategoryIndex)
                Case "A+": scoreList(studentIndex) = scoreList(studentIndex) + 5
                Case "A":  scoreList(studentIndex) = scoreList(studentIndex) + 4
                Case "B+": scoreList(studentIndex) = scoreList(studentIndex) + 3
                Case "B":  scoreList(studentIndex) = scoreList(studentIndex) + 2
                Case "C":  scoreList(studentIndex) = scoreList(studentIndex) + 1
            End Select
        Next gradeCategoryIndex
    Next studentIndex
End Sub

Private Sub BubbleSortTwoLists(ByRef nameList() As String, ByRef scoreList() As Double)
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim valueSwapped As Boolean

    n = UBound(scoreList)

    For i = 1 To n - 1
        valueSwapped = False

        For j = 1 To n - i
            If scoreList(j) < scoreList(j + 1) Then
                SwapPlaces nameList(j), nameList(j + 1)
                SwapPlaces scoreList(j), scoreList(j + 1)
                valueSwapped = True
            End If
        Next j

        If Not valueSwapped Then
            Exit For
        End If
    Next i
End Sub

Public Sub SwapPlaces(ByRef firstValue As Variant, ByRef secondValue As Variant)
    Dim temp As Variant
    
    temp = firstValue
    firstValue = secondValue
    secondValue = temp
End Sub

Private Sub PopulateWinnerSelections(ByRef ws As Worksheet, ByRef nameList() As String)
    Const WINNER_COL As String = "L"
    Dim nameToWrite  As String
    Dim i            As Long

    For i = 1 To UBound(nameList)
        If nameList(i) <> vbNullString Then
            ' Need to also update code to shade their names in the name list
            ' nameToWrite = TrimWinnersListName(nameList(i))
            WriteNewRangeValue ws.Range(WINNER_COL & (i + 1)), nameList(i)
        End If
    Next i
End Sub

Private Function TrimWinnersListName(ByVal nameToTrim As String) As String
    Dim engName As String
    Dim korName As String
    
    SplitWinnerName nameToTrim, engName, korName
    
    If Len(engName) > 25 Then
        engName = Left$(engName, 25)
    End If
    
    TrimWinnersListName = engName & " (" & korName & ")"
End Function