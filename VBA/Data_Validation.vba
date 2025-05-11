Option Explicit

#Const PRINT_DEBUG_MESSAGES = True
#If Mac Then
    Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
    Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
#End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data Validation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function VerifyRecordsAreComplete(ByVal ws As Worksheet, ByRef lastRow As Long, ByRef firstStudentRecord As Long) As Boolean
    Dim classInfoData As Variant
    Dim studentData As Variant
    Dim msgresult As Variant
    Dim classInfoStart As String
    Dim classInfoEnd As String
    Dim studentInfoStart As String
    Dim studentInfoEnd As String
    
    Const CLASS_INFO_COL As String = "C"
    Const CLASS_INFO_FIRST_ROW As Long = 1
    Const CLASS_INFO_LAST_ROW As Long = 6
    Const STUDENT_INFO_FIRST_ROW As Long = 8
    Const STUDENT_INFO_FIRST_COL As String = "B"
    Const STUDENT_INFO_LAST_COL As String = "K"
    Const MSG_NO_STUDENT_RECORDS As String = "No students were found!"
    Const MSG_ERROR_LOADING_DATA As String = "Could not read class or student data."
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Verifying Student Records Are Complete"
    #End If
    
    firstStudentRecord = STUDENT_INFO_FIRST_ROW ' Set here and passed back to keep things organized
    
    On Error Resume Next
    lastRow = ws.Cells.Item(ws.Rows.Count, STUDENT_INFO_FIRST_COL).End(xlUp).Row
    
    If lastRow < STUDENT_INFO_FIRST_ROW Then
        msgresult = DisplayMessage(MSG_NO_STUDENT_RECORDS, vbOKOnly + vbCritical, "Error!", 160)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    No students were found."
        #End If
        VerifyRecordsAreComplete = False
        Exit Function
    End If
    
    classInfoStart = CLASS_INFO_COL & CLASS_INFO_FIRST_ROW
    classInfoEnd = CLASS_INFO_COL & CLASS_INFO_LAST_ROW
    studentInfoStart = STUDENT_INFO_FIRST_COL & STUDENT_INFO_FIRST_ROW
    studentInfoEnd = STUDENT_INFO_LAST_COL & lastRow
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Final student record entry: " & (lastRow - STUDENT_INFO_FIRST_ROW + 1) & vbNewLine & _
                    "    Validating entered records"
    #End If
    
    classInfoData = ws.Range(classInfoStart & ":" & classInfoEnd).Value
    studentData = ws.Range(studentInfoStart & ":" & studentInfoEnd).Value ' Read in all student records
    On Error GoTo 0
    
    If IsEmpty(classInfoData) Or IsEmpty(studentData) Then
        msgresult = DisplayMessage(MSG_ERROR_LOADING_DATA, vbOKOnly + vbCritical, "Error!", 220)
        VerifyRecordsAreComplete = False
        Exit Function
    End If
    
    If Not ValidateClassInfoFromArray(ws, classInfoData) Or Not ValidateStudentInfoFromArray(ws, studentData) Then
        VerifyRecordsAreComplete = False
        Exit Function
    End If
    
    VerifyRecordsAreComplete = True
End Function

Private Function ValidateData(ByVal currentCell As Range, ByVal dataType As String) As Boolean
    Dim dataValue As String
    
    ' Static declarations to make subsequent runs more efficient
    Static validLevels As Variant
    Static validDays As Variant
    Static validTimes As Variant
    Static gradeMapping As Variant
    Static isDeclared As Boolean

    If Not isDeclared Then
        validLevels = Array("Theseus", "Perseus", "Odysseus", "Hercules", "Artemis", "Hermes", "Apollo", _
                            "Zeus", "E5 Athena", "Helios", "Poseidon", "Gaia", "Hera", "E6 Song's")
        validDays = Array("MonWed", "MonFri", "WedFri", "MWF", "TTh", "MWF (Class 1)", "MWF (Class 2)", _
                          "TTh (Class 1)", "TTh (Class 2)")
        validTimes = Array("9am", "10am", "11am", "12pm", "1pm", "2pm", "3pm", "4pm", "5pm", "530pm", _
                           "6pm", "7pm", "8pm", "830pm", "9pm")
        gradeMapping = Array("C", "B", "B+", "A", "A+")
        isDeclared = True
    End If
    
    dataValue = Trim$(currentCell.Value)
    
    Select Case dataType
        Case "Level:"
            ValidateData = IsValueValid(validLevels, dataValue)
        Case "Class Days:"
            ValidateData = IsValueValid(validDays, dataValue)
        Case "(Class 1) Time:"
            ValidateData = IsValueValid(validTimes, dataValue)
        Case "Grammar", "Pronunciation", "Fluency", "Manner", "Content", "Overall Effort"
            dataValue = UCase$(dataValue)
            If IsValueValid(gradeMapping, dataValue) Then
                currentCell.Value = dataValue
                ValidateData = True
            ElseIf IsNumeric(dataValue) And val(dataValue) >= 1 And val(dataValue) <= 5 Then
                ' Map a numeric value to its matching grade by its array index
                currentCell.Value = gradeMapping(val(dataValue) - 1)
                ValidateData = True
            Else
                ValidateData = False
            End If
        Case "Comments"
            ValidateData = (Len(dataValue) < 961)
        Case Else
            ValidateData = False
    End Select
End Function

Private Function IsValueValid(ByRef dataArray As Variant, ByVal dataValue As String) As Boolean
    Dim i As Long
    
    For i = LBound(dataArray) To UBound(dataArray)
        If dataArray(i) = dataValue Then
            IsValueValid = True
            Exit Function
        End If
    Next i
    IsValueValid = False
End Function

Private Function ValidateClassInfoFromArray(ByVal ws As Worksheet, ByVal classInfoData As Variant) As Boolean
    Dim cellValue As String
    Dim labelValue As String
    Dim msgToDisplay As String
    Dim msgresult As Long
    Dim i As Long

    ValidateClassInfoFromArray = True ' Assume valid initially

    With ws.Cells
        For i = 1 To UBound(classInfoData, 1) ' Iterate through rows of the array
            cellValue = Trim$(CStr(classInfoData(i, 1))) ' Data is in the first (only) column read
            labelValue = Trim$(.Item(i, 1).Value) ' Get label from sheet (less critical performance-wise)
    
            If cellValue = vbNullString Then
                msgToDisplay = "Class information incomplete." & vbNewLine & vbNewLine & _
                               "Missing: " & Left$(labelValue, Len(labelValue) - 1)
                msgresult = DisplayMessage(msgToDisplay, vbOKOnly, "Error!", 190)
                #If PRINT_DEBUG_MESSAGES Then
                    Debug.Print "Class information incomplete." & vbNewLine & _
                                "Missing: " & Left$(labelValue, Len(labelValue) - 1)
                #End If
                ValidateClassInfoFromArray = False
                Exit Function
            End If
    
            If i >= 3 And i <= 5 Then
                If Not ValidateData(.Item(i, 3), labelValue) Then ' Check if needs adapting
                    msgToDisplay = "Invalid value for " & Left$(.Item(i, 1).Value, Len(.Item(i, 1).Value) - 1) & "." & vbNewLine & vbNewLine & _
                                   "Would you like to ignore and continue?"
                    If DisplayMessage(msgToDisplay, vbYesNo, "Error!", 250) = vbNo Then
                        #If PRINT_DEBUG_MESSAGES Then
                            Debug.Print "    Invalid value for: " & Left$(.Item(i, 1).Value, Len(.Item(i, 1).Value) - 1)
                        #End If
                        ValidateClassInfoFromArray = False
                        Exit Function
                    End If
                End If
            End If
        Next i
    End With
End Function

Private Function ValidateStudentInfoFromArray(ByVal ws As Worksheet, ByVal studentData As Variant) As Boolean
    Dim rowNumber As Long
    Dim colNumber As Long
    Dim studentIndex As Long
    Dim cellValue As String
    Dim headerValue As String
    Dim studentEngName As String
    Dim msgToDisplay As String
    Dim dialogSize As Long
    Dim msgresult As Long

    ValidateStudentInfoFromArray = True ' Assume valid initially

    With ws.Cells
        For rowNumber = 1 To UBound(studentData, 1) ' Iterate rows (students)
            studentIndex = rowNumber + 7 ' Calculate original sheet row index
            studentEngName = Trim$(CStr(studentData(rowNumber, 1))) ' English Name (Col B = Index 1 in array)
    
            For colNumber = 1 To UBound(studentData, 2) ' Iterate columns (B to K = Index 1 to 10)
                cellValue = Trim$(CStr(studentData(rowNumber, colNumber)))
                headerValue = Trim$(.Item(7, colNumber + 1).Value) ' Get header from Row 7 (Column Index + 1)
    
                ' Skip validation for Korean Name (Col C / Index 2) and Notes (Col K / Index 10) for emptiness
                If colNumber <> 2 And colNumber <> 10 Then
                    If cellValue = vbNullString Then
                        msgToDisplay = "Student information incomplete." & vbNewLine & vbNewLine & _
                                       "Missing data for student " & studentIndex - 7 & "'s " & _
                                       IIf(colNumber = 1, headerValue, "(" & studentEngName & ") " & UCase$(headerValue)) & _
                                       IIf(colNumber = 9, " comment.", " score.") ' Col J = Index 9
                        dialogSize = 250
                        GoTo ErrorHandler
                    End If
                End If
    
                If colNumber >= 3 And colNumber <= 8 Then ' Grades (Cols D-I = Index 3 to 8)
                     If Not ValidateData(.Item(studentIndex, colNumber + 1), headerValue) Then
                        msgToDisplay = "Invalid value entered for student " & studentIndex - 7 & "'s (" & studentEngName & ") " & UCase$(headerValue) & " score."
                        dialogSize = 300
                        GoTo ErrorHandler
                    End If
                ElseIf colNumber = 9 Then ' Comment (Col J = Index 9)
                    If Not ValidateData(.Item(studentIndex, colNumber + 1), "Comments") Then ' Check if ValidateData handles length
                        msgToDisplay = "The COMMENT for student " & studentIndex - 7 & "'s (" & studentEngName & ") is too long. Please shorten it by " & _
                                       Len(.Item(rowNumber + 7, colNumber + 1).Value) - 960 & " or more characters."
                        dialogSize = 330
                        GoTo ErrorHandler
                    End If
                End If
            Next colNumber
        Next rowNumber
    End With

    Exit Function
ErrorHandler:
    msgresult = DisplayMessage(msgToDisplay, vbExclamation, "Error!", dialogSize)
    ValidateStudentInfoFromArray = False
End Function
