Option Explicit

#Const Windows = (Mac = 0)

Public Enum CellShading
    None = xlNone
    White = 16777215        ' RGB(255, 255, 255)
    LightPeach = 14281213   ' RGB(253, 233, 217)
    Orange = 9420794        ' RGB(250, 191, 143)
    LightGrey = 15921906    ' RGB(242, 242, 242)
    mediumGrey = 14277081   ' RGB(217, 217, 217)
    Grey = 12566463         ' RGB(191, 191, 191)
    LightPink = 14408946    ' RGB(242, 220, 219)
    DeepPink = 9737946      ' RGB(218, 150, 148)
    LightGreen = 14610923   ' RGB(235, 241, 222)
    Green = 10213316        ' RGB(196, 215, 155)
    Lavender = 15523812     ' RGB(228, 223, 236)
    Purple = 13082801       ' RGB(177, 160, 199)
    LightBlue = 15853276    ' RGB(220, 230, 241)
    SkyBlue = 15986394      ' RGB(218, 238, 243)
    MediumBlue = 14136213   ' RGB(149, 179, 215)
    Teal = 14470546         ' RGB(146, 205, 220)
    Beige = 12900829        ' RGB(221, 217, 196)
    Tan = 9944516           ' RGB(196, 189, 151)
    LightSteal = 15849925   ' RGB(197, 217, 241)
    Yellow = 65535          ' RGB(255, 255, 0)
    Red = 255               ' RGB(255, 0, 0)
    Gold = 49407            ' RGB(255, 215, 0)
    Silver = 12632256       ' RGB(192, 192, 192)
    Bronze = 3309517        ' RGB(205, 127, 50)
End Enum

Public Sub UpdateClassRecords(ByRef ws As Worksheet, ByVal targetRange As Range)
    Dim validatedValues()    As String
    Dim cellCategory()       As String
    Dim numberOfUpdatedCells As Long
    Dim updateShading        As Boolean

    numberOfUpdatedCells = targetRange.Cells.Count

    ReDim validatedValues(1 To numberOfUpdatedCells)
    ReDim cellCategory(1 To numberOfUpdatedCells)

    If ValidateEnteredValues(ws, targetRange, numberOfUpdatedCells, validatedValues(), cellCategory()) Then
        UpdateCellShading ws, targetRange, numberOfUpdatedCells, validatedValues(), cellCategory()
    End If
End Sub

' ========= Data Validation Functions ==========

Private Function ValidateEnteredValues(ByRef ws As Worksheet, ByVal targetRange As Range, ByVal numberOfCells As Long, ByRef validatedValues() As String, ByRef cellCategory() As String) As Boolean
    Dim currentCell         As Range
    Dim cellValue           As String
    Dim studentNameUpdate   As Boolean
    Dim cellIndex           As Long
    Dim studentIndex        As Long
    Dim updateShading       As Boolean

    For Each currentCell In targetRange
        cellIndex = cellIndex + 1
        cellValue = Trim$(CStr(currentCell.Value))
        cellCategory(cellIndex) = GetCellCategory(currentCell)
        validatedValues(cellIndex) = cellValue
        
        If validatedValues(cellIndex) <> vbNullString Then
            validatedValues(cellIndex) = ValidateValueByCellCategory(cellCategory(cellIndex), cellValue)
        End If

        Select Case cellCategory(cellIndex)
            Case g_ENGLISH_NAMES, g_KOREAN_NAMES
                updateShading = True
                studentNameUpdate = True
                studentIndex = currentCell.Row - 7 ' Offset to match array index
            Case g_COMMENTS, g_WINNER_NAMES
                updateShading = True
        End Select

        If validatedValues(cellIndex) <> cellValue Then
            WriteNewRangeValue currentCell, validatedValues(cellIndex)
        End If

        If studentNameUpdate Then
            If IsAWinningStudent(currentCell.Interior.Color) Then
                UpdateWinnerNames ws, currentCell
            End If
            UpdateWinnersValidationList ws, ws.Range(g_FULL_NAMES).Value
            studentNameUpdate = False
        End If
    Next currentCell

    ValidateEnteredValues = updateShading
End Function

Private Function ValidateValueByCellCategory(ByVal cellCategory As String, ByVal cellValue As String) As String
    Dim displayMsg As String
    
    Select Case cellCategory
        Case g_NATIVE_TEACHER, g_ENGLISH_NAMES
            ValidateValueByCellCategory = FormatName(cellValue)
        Case g_KOREAN_TEACHER, g_KOREAN_NAMES
            displayMsg = IIf(cellCategory = g_KOREAN_TEACHER, "Display.StudentRecords.InvalidKoreanTeacherName", "Display.StudentRecords.InvalidKoreanName")
            
            If IsValueEnglish(cellValue) Then
                DisplayMessage displayMsg
                ValidateValueByCellCategory = vbNullString
            Else
                ValidateValueByCellCategory = cellValue
            End If
        ' Case "C3" ' Class Level
        '     ValidateValueByCellCategory = ValidateClassLevel(cellValue)
        ' Case "C4" ' Class Days
        '     ValidateValueByCellCategory = ValidateClassDays(cellValue)
        ' Case "C5" ' Class Time
        '     ValidateValueByCellCategory = ValidateClassTime(cellValue)
        Case g_EVALUATION_DATE
            ValidateValueByCellCategory = FormatEvalDate(cellValue)
        Case g_STUDENT_GRADES
            ValidateValueByCellCategory = FormatGrade(cellValue)
        Case g_COMMENTS
            ValidateValueByCellCategory = FormatComment(cellValue)
        ' Case g_WINNER_NAMES
        '     ValidateValueByCellCategory = ValidateWinnerName(cellValue)
        Case Else
            ValidateValueByCellCategory = cellValue
    End Select
End Function

Private Function GetCellCategory(ByRef targetCell As Range) As String
    Dim studentRecordRanges() As Variant
    Dim i As Long

    studentRecordRanges() = Array(g_NATIVE_TEACHER, g_KOREAN_TEACHER, g_CLASS_LEVEL, g_CLASS_DAYS, g_CLASS_TIME, g_EVALUATION_DATE, _
                                  g_ENGLISH_NAMES, g_KOREAN_NAMES, g_STUDENT_GRADES, g_COMMENTS, g_TEACHER_NOTES, g_WINNER_NAMES)
    
    For i = LBound(studentRecordRanges) To UBound(studentRecordRanges)
        If Not Intersect(targetCell, targetCell.Worksheet.Range(studentRecordRanges(i))) Is Nothing Then
            GetCellCategory = studentRecordRanges(i)
            Exit Function
        End If
    Next i

    GetCellCategory = "Unknown"
End Function

Private Sub UpdateWinnersValidationList(ByRef ws As Worksheet, ByRef nameList As Variant)
    Dim validationList As String
    Dim mergedName As String
    Dim i As Long

    For i = LBound(nameList) To UBound(nameList)
        If nameList(i, 1) <> vbNullString And nameList(i, 2) <> vbNullString Then
            mergedName = GetMergedName(nameList(i, 1), nameList(i, 2))
            validationList = validationList & mergedName & ","
        End If
    Next i

    If Right$(validationList, 1) = "," Then
        validationList = Left$(validationList, Len(validationList) - 1)
    End If

        If validationList <> vbNullString Then
        With ws
            On Error Resume Next
            .Range(g_WINNER_NAMES).Validation.Delete
            On Error GoTo 0
    
            With .Range(g_WINNER_NAMES).Validation
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=validationList
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
        End With
    End If
End Sub

Private Sub UpdateWinnerNames(ByRef ws As Worksheet, ByRef currentCell As Range)
    Dim targetCell  As Range
    Dim engName     As String
    Dim korName     As String

    With currentCell
        Select Case .Interior.Color
            Case CellShading.Gold
                Set targetCell = ws.Range("L2")
            Case CellShading.Silver
                Set targetCell = ws.Range("L3")
            Case CellShading.Bronze
                Set targetCell = ws.Range("L4")
            Case Else
                Exit Sub
        End Select

        Select Case .Column
            Case 2
                targetCell.Value = GetMergedName(.Value, .Offset(0, 1).Value)
            Case 3
                targetCell.Value = GetMergedName(.Offset(0, -1).Value, .Value)
            Case Else
                Exit Sub
        End Select
    End With
End Sub

' ========= Shading Functions ==========

Private Function IsAWinningStudent(ByVal cellColor As Long) As Boolean
    Select Case cellColor
        Case CellShading.Gold, CellShading.Silver, CellShading.Bronze
            IsAWinningStudent = True
        Case Else
            IsAWinningStudent = False
    End Select
End Function

Public Sub SetDefaultShading(ByRef ws As Worksheet)
    Dim dicShadingUpdates   As New Dictionary
    Dim currentCell         As Range
    Dim rngWinnerNames      As Range
    Dim rangeToShade        As Range
    Dim rngEnglishNames     As Range
    Dim rngKoreanNames      As Range
    Dim rngComments         As Range
    Dim englishName         As String
    Dim koreanName          As String
    Dim shadingValue        As Long
    Dim winnerPlacement     As Long
    Dim valueLength         As Long
    Dim i                   As Long

    With ws
        Set rngEnglishNames = .Range(g_ENGLISH_NAMES)
        Set rngKoreanNames = .Range(g_KOREAN_NAMES)
        Set rngComments = .Range(g_COMMENTS)
        Set rngWinnerNames = .Range(g_WINNER_NAMES)
    End With

    For Each currentCell In rngEnglishNames
        valueLength = Len(currentCell.Value)
        AddValueToDictionary dicShadingUpdates, currentCell.Address, GetEnglishNameShading(valueLength, False)
    Next currentCell

    For Each currentCell In rngKoreanNames
        valueLength = Len(currentCell.Value)
        AddValueToDictionary dicShadingUpdates, currentCell.Address, GetKoreanNameShading(valueLength, False)
    Next currentCell

    For Each currentCell In rngComments
        valueLength = Len(currentCell.Value)
        AddValueToDictionary dicShadingUpdates, currentCell.Address, GetCommentShading(valueLength, False)
    Next currentCell

    For Each currentCell In rngWinnerNames
        winnerPlacement = winnerPlacement + 1

        shadingValue = GetShadingForWinners(winnerPlacement)
        SplitWinnerName currentCell.Value, englishName, koreanName

        If englishName <> vbNullString And koreanName <> vbNullString Then
            Set rangeToShade = FindNameInStudentList(ws, englishName, koreanName)
            If Not rangeToShade Is Nothing Then
                AddValueToDictionary dicShadingUpdates, rangeToShade.Address & ":" & rangeToShade.Offset(0, 1).Address, shadingValue
            End If
        End If

        englishName = vbNullString
        koreanName = vbNullString
    Next currentCell

    ApplyShading ws, dicShadingUpdates
End Sub

Public Function GetEnglishNameShading(ByVal nameLength As Long, Optional ByVal enableWarningMessage As Boolean = True) As Long
    Select Case nameLength
        Case 0 To 21
            GetEnglishNameShading = CellShading.White
        Case Else
            If enableWarningMessage Then
                DisplayMessage "Display.StudentRecords.EnglishNameTooLongForReport"
            End If
            GetEnglishNameShading = CellShading.Red
    End Select
End Function

Public Function GetKoreanNameShading(ByVal nameLength As Long, Optional ByVal enableWarningMessage As Boolean = True) As Long
    Select Case nameLength
        Case 0, 3
            GetKoreanNameShading = CellShading.White
        Case 2, 4
            If enableWarningMessage Then
                DisplayMessage "Display.StudentRecords.KoreanNameUncommonLength", nameLength
            End If
            GetKoreanNameShading = CellShading.Yellow
        Case Else
            If enableWarningMessage Then
                DisplayMessage "Display.StudentRecords.KoreanNameInvalidLength"
            End If
            GetKoreanNameShading = CellShading.Red
    End Select
End Function

Public Function GetCommentShading(ByVal commentLength As Long, Optional ByVal enableWarningMessage As Boolean = True) As Long
    Const MIN_LEN As Long = 80
    Const MAX_LEN As Long = 960
    
    Select Case commentLength
        Case 1 To MIN_LEN - 1 ' Comment is too short
            If enableWarningMessage Then
                DisplayMessage "Display.StudentRecords.CommentTooShortUponEntry"
            End If
            GetCommentShading = CellShading.Yellow
        Case Is > MAX_LEN ' Comment exceeds max length
            If enableWarningMessage Then
                DisplayMessage "Display.StudentRecords.CommentTooLongUponEntry", CStr(commentLength), CStr(commentLength - MAX_LEN)
            End If
            GetCommentShading = CellShading.Red
        Case Else
            GetCommentShading = CellShading.LightGrey
    End Select
End Function

Private Function GetShadingForWinners(ByVal winnerPlacement As Long) As Long
    Select Case winnerPlacement
        Case 1
            GetShadingForWinners = CellShading.Gold
        Case 2
            GetShadingForWinners = CellShading.Silver
        Case 3
            GetShadingForWinners = CellShading.Bronze
    End Select
End Function

Public Sub SplitWinnerName(ByVal fullName As String, ByRef englishName As String, ByRef koreanName As String)
    Const DELIMITER As String = " ("
    Dim nameParts() As String
    
    If fullName <> vbNullString Then
        nameParts = Split(fullName, DELIMITER)
        englishName = Trim$(nameParts(0))
        koreanName = Trim$(Left$(nameParts(1), Len(nameParts(1)) - 1)) ' Remove closing parenthesis
    End If
End Sub

Private Function FindNameInStudentList(ByRef ws As Worksheet, ByVal englishName As String, ByVal koreanName As String) As Range
    Dim nameList As Variant
    Dim i As Long
    
    nameList = ws.Range(g_FULL_NAMES)

    For i = LBound(nameList) To UBound(nameList)
        If StrComp(nameList(i, 1), englishName, vbTextCompare) = 0 Then
            If StrComp(nameList(i, 2), koreanName, vbTextCompare) = 0 Then
                Set FindNameInStudentList = ws.Cells(i + 7, 2) ' Offset by 7 rows to match actual worksheet row
                Exit Function
            End If
        End If
    Next i

    Set FindNameInStudentList = Nothing
End Function

Public Sub ApplyShading(ByRef ws As Worksheet, ByRef dicShadingUpdates As Dictionary)
    Dim Key As Variant

    For Each Key In dicShadingUpdates.Keys
        ws.Range(Key).Interior.Color = dicShadingUpdates(Key)
    Next Key
End Sub

Private Sub UpdateCellShading(ByRef ws As Worksheet, ByVal targetRange As Range, ByVal numberOfCellsUpdated As Long, ByRef validatedValues() As String, ByRef cellCategory() As String)
    Dim dicShadingUpdates  As New Dictionary
    Dim currentCell        As Range
    Dim notInWinnersList   As Boolean
    Dim winnersListUpdated As Boolean
    Dim i                  As Long

    For Each currentCell In targetRange
        i = i + 1

        ' First determine if the current cell is a name cell that is not in the winners list
        If cellCategory(i) = g_ENGLISH_NAMES Or cellCategory(i) = g_KOREAN_NAMES Then
            notInWinnersList = Not IsAWinningStudent(currentCell.Interior.Color)
        End If

        ' Next, check if the winners list has been updated and determine necessary shading updates
        winnersListUpdated = (Not Intersect(currentCell, ws.Range(g_WINNER_NAMES)) Is Nothing)
        If winnersListUpdated Then
            UpdateWinnerStatusShading ws, dicShadingUpdates, currentCell.Address, validatedValues(i)
        End If

        ' Finally, prepare shading updates for non-winning-name and comment cell updates
        If notInWinnersList Or cellCategory(i) = g_COMMENTS Then
            PrepareUpdatedShadingForNamesAndComments dicShadingUpdates, currentCell, validatedValues(i), cellCategory(i)
        End If

        ' Reset flag for next iteration
        notInWinnersList = False
    Next currentCell

    ApplyShading ws, dicShadingUpdates
End Sub

Private Sub UpdateWinnerStatusShading(ByRef ws As Worksheet, ByRef dicShadingUpdates As Dictionary, ByVal cellAddress As String, ByRef validatedValue As String)
    Dim shadingValue     As Long

    Select Case cellAddress
        Case "$L$2"
            shadingValue = CellShading.Gold
        Case "$L$3"
            shadingValue = CellShading.Silver
        Case "$L$4"
            shadingValue = CellShading.Bronze
    End Select

    If validatedValue <> vbNullString Then
        SetShadingForWinnerName ws, validatedValue, dicShadingUpdates, shadingValue
        RemoveDuplicateWinners ws, cellAddress, validatedValue
    End If

    DetermineNameCellShading ws, dicShadingUpdates, ws.Range(g_ENGLISH_NAMES), shadingValue
End Sub

Private Sub SetShadingForWinnerName(ByRef ws As Worksheet, ByVal validatedValue As String, ByRef dicShadingUpdates As Dictionary, ByVal shadingValue As Long)
    Dim studentToFind    As Range
    Dim cellAddrToShade  As String
    Dim englishName      As String
    Dim koreanName       As String

    SplitWinnerName validatedValue, englishName, koreanName
    Set studentToFind = FindNameInStudentList(ws, englishName, koreanName)

    If Not studentToFind Is Nothing Then
        cellAddrToShade = studentToFind.Address & ":" & studentToFind.Offset(0, 1).Address
        AddValueToDictionary dicShadingUpdates, cellAddrToShade, shadingValue
    End If
End Sub

Private Sub RemoveDuplicateWinners(ByRef ws As Worksheet, ByVal cellAddress As String, ByVal validatedValue As String)
    Dim currentCell As Range

    For Each currentCell In ws.Range(g_WINNER_NAMES)
        With currentCell
            If .Address <> cellAddress Then
                If StrComp(.Value, validatedValue, vbTextCompare) = 0 Then
                    .Value = vbNullString
                End If
            End If
        End With
    Next currentCell
End Sub

Private Sub DetermineNameCellShading(ByRef ws As Worksheet, ByRef dicShadingUpdates As Dictionary, ByVal engNamesRange As Range, ByVal shadingValue As Long)
    Dim currentCell As Range
    Dim valueLength As Long

    For Each currentCell In engNamesRange
        If currentCell.Interior.Color = shadingValue Then
            valueLength = Len(currentCell.Value)
            AddValueToDictionary dicShadingUpdates, currentCell.Address, GetEnglishNameShading(valueLength, False)

            valueLength = Len(currentCell.Offset(0, 1).Value)
            AddValueToDictionary dicShadingUpdates, currentCell.Offset(0, 1).Address, GetKoreanNameShading(valueLength, False)
            
            Exit For
        End If
    Next currentCell
End Sub

Private Sub PrepareUpdatedShadingForNamesAndComments(ByRef dicShadingUpdates As Dictionary, ByVal currentCell As Range, ByRef validatedValue As String, ByRef cellCategory As String)
    Dim newShadingValue         As Long
    Dim valueLength             As Long
    Dim clearShading            As Boolean
    Dim updateShadingDictionary As Boolean
    Dim otherHalfOfName         As String

    Select Case cellCategory
        Case g_ENGLISH_NAMES
            updateShadingDictionary = True

            If validatedValue = vbNullString Then
                clearShading = IsAWinningStudent(currentCell.Interior.Color)
                newShadingValue = CellShading.None
            Else
                valueLength = Len(validatedValue)
                newShadingValue = GetEnglishNameShading(valueLength)
            End If
        Case g_KOREAN_NAMES
            updateShadingDictionary = True

            If validatedValue = vbNullString Then
                clearShading = IsAWinningStudent(currentCell.Interior.Color)
                newShadingValue = CellShading.None
            Else
                valueLength = Len(validatedValue)
                newShadingValue = GetKoreanNameShading(valueLength)
            End If
        Case g_COMMENTS
            updateShadingDictionary = True

            If validatedValue = vbNullString Then
                newShadingValue = CellShading.None
            Else
                valueLength = Len(validatedValue)
                newShadingValue = GetCommentShading(valueLength)
            End If
    End Select

    If updateShadingDictionary Then AddValueToDictionary dicShadingUpdates, currentCell.Address, newShadingValue

    If clearShading Then
        If cellCategory = g_ENGLISH_NAMES Then
            otherHalfOfName = currentCell.Offset(0, 1).Address
            valueLength = Len(currentCell.Offset(0, 1).Value)
            newShadingValue = GetKoreanNameShading(valueLength)
        ElseIf cellCategory = g_KOREAN_NAMES Then
            otherHalfOfName = currentCell.Offset(0, -1).Address
            valueLength = Len(currentCell.Offset(0, -1).Value)
            newShadingValue = GetEnglishNameShading(valueLength)
        End If

        AddValueToDictionary dicShadingUpdates, otherHalfOfName, newShadingValue
    End If
End Sub