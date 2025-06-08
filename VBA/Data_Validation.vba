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
    Dim msgResult As Variant
    Dim classInfoLabels As String
    Dim classInfoEnd As String
    Dim studentInfoStart As String
    Dim studentInfoEnd As String
    
    Const CLASS_INFO_RANGE As String = "A1:C6"
    Const STUDENT_INFO_FIRST_ROW As Long = 8
    Const STUDENT_INFO_FIRST_COL As String = "B"
    Const STUDENT_INFO_LAST_COL As String = "K"
    Const MSG_NO_STUDENT_RECORDS As String = "No students were found!"
    Const MSG_ERROR_LOADING_DATA As String = "Could not read class or student data."
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Verifying Student Records Are Complete"
    #End If
    
    On Error Resume Next
    lastRow = ws.Cells.Item(ws.Rows.Count, STUDENT_INFO_FIRST_COL).End(xlUp).Row
    firstStudentRecord = STUDENT_INFO_FIRST_ROW ' Set here and passed back to keep things organized
    studentInfoStart = STUDENT_INFO_FIRST_COL & STUDENT_INFO_FIRST_ROW
    studentInfoEnd = STUDENT_INFO_LAST_COL & lastRow
    
    If lastRow < STUDENT_INFO_FIRST_ROW Then
        msgResult = DisplayMessage(MSG_NO_STUDENT_RECORDS, vbOKOnly + vbCritical, "Error!", 160)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print INDENT_LEVEL_1 & "No students were found."
        #End If
        VerifyRecordsAreComplete = False
        Exit Function
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print INDENT_LEVEL_1 & "Final student record entry: " & (lastRow - STUDENT_INFO_FIRST_ROW + 1) & vbNewLine & _
                    INDENT_LEVEL_1 & "Validating entered records"
    #End If
    
    classInfoData = ws.Range(CLASS_INFO_RANGE).Value
    studentData = ws.Range(studentInfoStart & ":" & studentInfoEnd).Value ' Read in all student records
    
    On Error GoTo 0
    If IsEmpty(classInfoData) Or IsEmpty(studentData) Then
        msgResult = DisplayMessage(MSG_ERROR_LOADING_DATA, vbOKOnly + vbCritical, "Error!", 220)
        VerifyRecordsAreComplete = False
        Exit Function
    End If
    
    If Not ValidateClassInfoFromArray(ws, classInfoData) Or Not ValidateStudentInfoFromArray(ws, studentData) Then
        VerifyRecordsAreComplete = False
        Exit Function
    End If
    
    VerifyRecordsAreComplete = True
End Function

Private Function ValidateData(ByVal dataValue As String, ByVal dataType As String, Optional ByVal targetCell As Range) As Boolean
    
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
    
    Select Case dataType
        Case "English Name"
            ValidateData = (Len(dataValue) < 41)
        Case "Level:"
            ValidateData = IsValueValid(validLevels, dataValue)
        Case "Class Days:"
            ValidateData = IsValueValid(validDays, dataValue)
        Case "(Class 1) Time:"
            ValidateData = IsValueValid(validTimes, dataValue)
        Case "Grammar", "Pronunciation", "Fluency", "Manner", "Content", "Overall Effort"
            dataValue = UCase$(dataValue)
            If IsValueValid(gradeMapping, dataValue) Then
                If targetCell.Value <> dataValue Then targetCell.Value = dataValue
                ValidateData = True
            ElseIf IsNumeric(dataValue) And Val(dataValue) >= 1 And Val(dataValue) <= 5 Then
                ' Map a numeric value to its matching grade by its array index
                If targetCell.Value <> gradeMapping(Val(dataValue) - 1) Then targetCell.Value = gradeMapping(Val(dataValue) - 1)
                ValidateData = True
            Else
                ValidateData = False
            End If
        Case "Comments (Positive - Negative - Positive)"
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
    Dim dataCategory As Long
    Dim dialogSize As Long

    ValidateClassInfoFromArray = True ' Assume valid initially

    ' Skip index 2 as it contains no data
    For dataCategory = LBound(classInfoData, 1) To UBound(classInfoData, 1)
        labelValue = Trim$(CStr(classInfoData(dataCategory, 1)))
        cellValue = Trim$(CStr(classInfoData(dataCategory, 3)))

        If cellValue = vbNullString Then
            msgToDisplay = "Class information incomplete." & vbNewLine & vbNewLine & _
                           "Missing: " & Left$(labelValue, Len(labelValue) - 1)
            dialogSize = 190

            Call DisplayMessage(msgToDisplay, vbOKOnly, "Error!", dialogSize)

            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "Class information incomplete." & vbNewLine & _
                            "Missing: " & Left$(labelValue, Len(labelValue) - 1)
            #End If

            ValidateClassInfoFromArray = False
            Exit Function
        End If

        If dataCategory >= 3 And dataCategory <= 5 Then
            If Not ValidateData(cellValue, labelValue) Then
                msgToDisplay = "Invalid value for " & Left$(labelValue, Len(labelValue) - 1) & "." & vbNewLine & vbNewLine & _
                                "Would you like to ignore and continue?"
                dialogSize = 250

                If DisplayMessage(msgToDisplay, vbYesNo, "Error!", dialogSize) = vbNo Then
                    #If PRINT_DEBUG_MESSAGES Then
                        Debug.Print INDENT_LEVEL_1 & "Invalid value for: " & Left$(labelValue, Len(labelValue) - 1)
                    #End If
                    ValidateClassInfoFromArray = False
                    Exit Function
                End If
            End If
        End If
    Next dataCategory
End Function

Private Function ValidateStudentInfoFromArray(ByVal ws As Worksheet, ByVal studentData As Variant) As Boolean
    Dim targetCell As Range
    Dim englishName As String
    Dim koreanName As String
    Dim studentName As String
    Dim categoryName As String
    Dim msgToDisplay As String
    Dim currentStudent As Long
    Dim dataCategory As Long
    Dim dialogSize As Long

    Const HEADER_ROW As Long = 7
    Const DATA_CATEGORIES As Long = 9
    Const ENGLISH_NAME_COLUMN As Long = 1
    Const KOREAN_NAME_COLUMN As Long = 2
    Const GRADE_COLUMNS_START As Long = 3
    Const GRADE_COLUMNS_END As Long = 8
    Const COMMENT_COLUMN As Long = 9
    Const MAX_ENGLISH_NAME_LENGTH As Long = 40
    Const MAX_COMMENT_LENGTH As Long = 960

    ' Assume data is valid by default
    ValidateStudentInfoFromArray = True

    ' Reminder: 1 and 2 refer to the array's dimensions
    For currentStudent = LBound(studentData, 1) To UBound(studentData, 1)
        ' Set student name for error and debug messages
        englishName = Trim$(CStr(studentData(currentStudent, ENGLISH_NAME_COLUMN)))
        koreanName = Trim$(CStr(studentData(currentStudent, KOREAN_NAME_COLUMN)))

        Select Case True
            Case englishName <> vbNullString And koreanName <> vbNullString
                studentName = englishName & "(" & koreanName & ")"
            Case englishName = vbNullString
                studentName = koreanName & " (Row: " & currentStudent & ")"
                msgToDisplay = "No value has been entered for " & studentName & " 's English name. Please enter a " & _
                               "value and try again."
                dialogSize = 250
                GoTo ErrorHandler
            Case koreanName = vbNullString
                studentName = englishName & " (Row: " & currentStudent & ")"
                msgToDisplay = "No value has been entered for " & studentName & " 's Korean name. Please enter a " & _
                               "value and try again."
                dialogSize = 250
                GoTo ErrorHandler
        End Select
        
        For dataCategory = GRADE_COLUMNS_START To COMMENT_COLUMN
            Set targetCell = ws.Cells.Item(HEADER_ROW + currentStudent, 1 + dataCategory)
            
            ' Set category name for data validation and debug & error messages
            categoryName = Trim$(ws.Cells.Item(HEADER_ROW, 1 + dataCategory).Value)

            ' Verify data has been entered
            If IsMissingData(studentData(currentStudent, dataCategory), studentName, categoryName) Then GoTo ErrorHandler

            If Not ValidateData(studentData(currentStudent, dataCategory), categoryName, targetCell) Then
                Select Case dataCategory
                    Case ENGLISH_NAME_COLUMN
                        msgToDisplay = studentName & "'s " & UCase$(categoryName) & " is too long. The maximum supported length is " & _
                                       CStr(MAX_ENGLISH_NAME_LENGTH) & " characters. Please shorten it by at least " & _
                                       Len(studentData(currentStudent, dataCategory)) - MAX_ENGLISH_NAME_LENGTH & " characters."
                        dialogSize = 250
                    Case GRADE_COLUMNS_START To GRADE_COLUMNS_END
                        msgToDisplay = "Invalid value entered for " & studentName & "'s " & UCase$(categoryName) & " score."
                        dialogSize = 300
                    Case COMMENT_COLUMN
                        msgToDisplay = "The COMMENT for " & studentName & " is too long. Please shorten it by " & _
                                       Len(studentData(currentStudent, dataCategory)) - 960 & " or more characters."
                        dialogSize = 330
                End Select

                GoTo ErrorHandler
            End If
        Next dataCategory
    Next currentStudent

    Exit Function
    
ErrorHandler:
    Call DisplayMessage(msgToDisplay, vbExclamation, "Error!", dialogSize)
    ValidateStudentInfoFromArray = False
End Function

Private Function IsMissingData(ByVal dataValue As Variant, ByVal studentName As String, ByVal categoryName As String) As Boolean
    Dim msgToDisplay As String
    Dim dialogSize As Long
    
    IsMissingData = False

    If dataValue = vbNullString Then
        msgToDisplay = "Student information incomplete." & vbNewLine & vbNewLine & _
                       "Missing data for " & studentName & "'s " & UCase$(categoryName) & "."
        dialogSize = 250

        Call DisplayMessage(msgToDisplay, vbExclamation, "Error!", dialogSize)
        IsMissingData = True
    End If
End Function
