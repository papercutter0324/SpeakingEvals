Option Explicit

#Const PRINT_DEBUG_MESSAGES = True
#If Mac Then
    Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
    Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
#End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Winners List Name Sorting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AutoSelectClassWinners(ByVal ws As Worksheet)
    Dim validationRange As Range
    Dim winnersRange As Range
    Dim nameList() As String
    Dim studentScores() As Double
    
    ' Set ranges
    Set validationRange = ws.Range("BB1:BB25")
    Set winnersRange = ws.Range("L2:L4")
    
    ' Initialize student names array
    ReDim nameList(0 To 24)
    GenerateNameListToSort ws, nameList()
    
    ' Calculate scores using an efficient subroutine
    ReDim studentScores(0 To UBound(nameList()))
    CalculateScores ws, studentScores()
    
    ' Sort students by scores
    SortStudentsByScoresBubble nameList(), studentScores()
    
    ' Populate winners range with validation
    PopulateWinnersRange ws, nameList()
    
    ' Validate and update winners names
    ws.Unprotect
    SetDefaultShading ws
    ws.Protect
End Sub

Private Sub GenerateNameListToSort(ByVal ws As Worksheet, ByRef nameList() As String)
    Dim rowIndex As Long
    Dim nameIndex As Long
    
    nameIndex = 0

    With ws
        For rowIndex = 8 To 32
            If Not IsEmpty(.Range("B" & rowIndex)) And Not IsEmpty(.Range("C" & rowIndex)) Then
                nameList(nameIndex) = .Range("B" & rowIndex).Value & "(" & .Range("C" & rowIndex).Value & ")"
                nameIndex = nameIndex + 1
            End If
        Next rowIndex
    End With
    
    ReDim Preserve nameList(nameIndex - 1)
End Sub

Private Sub CalculateScores(ByVal ws As Worksheet, ByRef studentScores() As Double)
    Dim studentsScoreRange As Range
    Dim studentIndex As Long
    Dim gradeCategoryIndex As Long
    
    Set studentsScoreRange = ws.Range("D8:I26")
    
    For studentIndex = 0 To UBound(studentScores())
        studentScores(studentIndex) = 0
        For gradeCategoryIndex = 1 To 6
            Select Case studentsScoreRange.Cells.Item(studentIndex, gradeCategoryIndex).Value
                Case "A+": studentScores(studentIndex) = studentScores(studentIndex) + 5
                Case "A":  studentScores(studentIndex) = studentScores(studentIndex) + 4
                Case "B+": studentScores(studentIndex) = studentScores(studentIndex) + 3
                Case "B":  studentScores(studentIndex) = studentScores(studentIndex) + 2
                Case "C":  studentScores(studentIndex) = studentScores(studentIndex) + 1
            End Select
        Next gradeCategoryIndex
    Next studentIndex
End Sub

Private Sub SortStudentsByScoresBubble(ByRef nameList() As String, ByRef studentScores() As Double)
    Dim i As Long
    Dim j As Long
    Dim n As Long
    
    n = UBound(nameList())
    
    For i = 1 To n - 1
        For j = 1 To n - i
            If studentScores(j) < studentScores(j + 1) Then
                SwapPlaces nameList(j), nameList(j + 1)
                SwapPlaces studentScores(j), studentScores(j + 1)
            End If
        Next j
    Next i
End Sub

Private Sub SwapPlaces(ByRef a As Variant, ByRef b As Variant)
    Dim temp As Variant
    
    temp = a
    a = b
    b = temp
End Sub

Private Sub PopulateWinnersRange(ByVal ws As Worksheet, ByRef nameList() As String)
    Dim i As Long
    
    For i = 2 To 4
        ws.Range("L" & i).Value = nameList(i - 2)
    Next i
End Sub
