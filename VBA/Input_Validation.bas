Option Explicit

#Const Windows = (Mac = 0)

Public Function ValidateStudentData(ByRef fullClassData() As StudentRecords, ByRef recordNumber As Long, ByRef invalidCategory As String) As Boolean
    Dim i As Long

    For i = LBound(fullClassData) To UBound(fullClassData)
        With fullClassData(i)
            If Not IsEnglishNameValid(.englishName) Then
                ' Display error
                recordNumber = i
                invalidCategory = "English Name"
                ValidateStudentData = False
                Exit Function
            End If

            If Not IsKoreanNameValid(.koreanName) Then
                ' Display error
                recordNumber = i
                invalidCategory = "Korean Name"
                ValidateStudentData = False
                Exit Function
            End If

            If Not IsScoreValid(.GrammarScore) Then
                recordNumber = i
                invalidCategory = "Grammar Score"
                ValidateStudentData = False
                Exit Function
            End If

            If Not IsScoreValid(.PronunciationScore) Then
                recordNumber = i
                invalidCategory = "Pronunciation Score"
                ValidateStudentData = False
                Exit Function
            End If

            If Not IsScoreValid(.FluencyScore) Then
                recordNumber = i
                invalidCategory = "Fluency Score"
                ValidateStudentData = False
                Exit Function
            End If

            If Not IsScoreValid(.MannerScore) Then
                recordNumber = i
                invalidCategory = "Manner Score"
                ValidateStudentData = False
                Exit Function
            End If

            If Not IsScoreValid(.ContentScore) Then
                recordNumber = i
                invalidCategory = "Content Score"
                ValidateStudentData = False
                Exit Function
            End If

            If Not IsScoreValid(.EffortScore) Then
                recordNumber = i
                invalidCategory = "Effort Score"
                ValidateStudentData = False
                Exit Function
            End If

            .OverallGrade = CalculateOverallGrade(fullClassData(i))

            If Not IsCommentValid(.comment) Then
                recordNumber = i
                invalidCategory = "Comment"
                ValidateStudentData = False
                Exit Function
            End If
        End With
    Next i

    ValidateStudentData = True
End Function

Private Function IsEnglishNameValid(ByVal englishName As String) As Boolean
    IsEnglishNameValid = (Len(englishName) <= 40)
End Function

Private Function IsKoreanNameValid(ByVal koreanName As String) As Boolean
    IsKoreanNameValid = (Len(koreanName) <= 5)
End Function

Private Function IsScoreValid(ByVal scoreToValidate As String) As Boolean
    Dim validScores As Variant
    Dim scoreValue  As String
    Dim isValid     As Boolean
    Dim tmpValue    As Long
    Dim i           As Long

    If IsNumeric(scoreToValidate) Then
        tmpValue = CLng(scoreToValidate)
        IsScoreValid = (tmpValue >= 1 And tmpValue <= 5)
        Exit Function
    End If
    
    scoreValue = UCase$(scoreToValidate)
    validScores = Array("C", "B", "B+", "A", "A+")

    For i = LBound(validScores) To UBound(validScores)
        If scoreValue = validScores(i) Then
            isValid = True
            Exit For
        End If
    Next i
    
    IsScoreValid = isValid
End Function

Private Function IsCommentValid(ByVal comment As String) As Boolean
    IsCommentValid = (Len(comment) <= 960)
End Function

Public Function ValidateClassData(ByRef classInformation As ClassRecords, ByRef invalidCategory As String) As Boolean
    With classInformation
        If .englishTeacher = vbNullString Then
            invalidCategory = "English Teacher"
            ValidateClassData = False
            Exit Function
        End If

        If .KoreanTeacher = vbNullString Then
            invalidCategory = "Korean Teacher"
            ValidateClassData = False
            Exit Function
        End If

        If Not IsClassLevelValid(.classLevel) Then
            invalidCategory = "Class Level"
            ValidateClassData = False
            Exit Function
        End If

        If Not IsClassDaysValid(.classDays) Then
            invalidCategory = "Class Days"
            ValidateClassData = False
            Exit Function
        End If

        If Not IsClassTimeValid(.classTime) Then
            invalidCategory = "Class Time"
            ValidateClassData = False
            Exit Function
        End If

        If Not IsEvaluationDateValid(.EvalulationDate) Then
            invalidCategory = "Evaluation Date"
            ValidateClassData = False
            Exit Function
        End If
    End With

    ValidateClassData = True
End Function

Private Function IsClassLevelValid(ByVal classLevel As String) As Boolean
    Dim validLevels As Variant
    Dim isValid As Boolean
    Dim i As Long

    If classLevel = vbNullString Then
        IsClassLevelValid = False
        Exit Function
    End If

    validLevels = Array("Theseus", "Perseus", "Odysseus", "Hercules", "Artemis", "Hermes", "Apollo", _
                        "Zeus", "E5 Athena", "Helios", "Poseidon", "Gaia", "Hera", "E6 Song's")

    For i = LBound(validLevels) To UBound(validLevels)
        If classLevel = validLevels(i) Then
            isValid = True
            Exit For
        End If
    Next i

    IsClassLevelValid = isValid
End Function

Private Function IsClassDaysValid(ByVal classDays As String) As Boolean
    Dim validDays As Variant
    Dim isValid As Boolean
    Dim i As Long

    If classDays = vbNullString Then
        IsClassDaysValid = False
        Exit Function
    End If

    validDays = Array("MonWed", "MonFri", "WedFri", "MWF", "TTh", "MWF (Class A)", _
                      "MWF (Class B)", "TTh (Class A)", "TTh (Class B)")

    For i = LBound(validDays) To UBound(validDays)
        If classDays = validDays(i) Then
            isValid = True
            Exit For
        End If
    Next i

    IsClassDaysValid = isValid
End Function

Private Function IsClassTimeValid(ByVal classTime As String) As Boolean
    Dim validTimes As Variant
    Dim isValid As Boolean
    Dim i As Long

    If classTime = vbNullString Then
        IsClassTimeValid = False
        Exit Function
    End If
    
    validTimes = Array("9am", "10am", "11am", "12pm", "1pm", "2pm", "3pm", "4pm", _
                       "5pm", "530pm", "6pm", "7pm", "8pm", "830pm", "9pm")

    For i = LBound(validTimes) To UBound(validTimes)
        If classTime = validTimes(i) Then
            isValid = True
            Exit For
        End If
    Next i

    IsClassTimeValid = isValid
End Function

Private Function IsEvaluationDateValid(ByVal evalDate As String) As Boolean
    IsEvaluationDateValid = IsDate(evalDate)
End Function

Public Function SheetContainsStudentRecords(ByVal finalRow As Long) As Boolean
    Const FIRST_STUDENT_ROW  As Long = 8

    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.StudentRecords.VerifyingIfComplete")
    End If
    
    SheetContainsStudentRecords = (finalRow >= FIRST_STUDENT_ROW)
End Function

Public Function VerifyRecordsAreComplete(ByRef ws As Worksheet, ByVal finalRow As Long) As Boolean
    Dim blankCount As Double
    Dim classInfo As Range
    Dim studentInfo As Range
    
    Set classInfo = ws.Range(g_CLASS_INFO)
    Set studentInfo = ws.Range("B8:J" & CStr(finalRow))

    blankCount = WorksheetFunction.CountBlank(classInfo)
    blankCount = blankCount + WorksheetFunction.CountBlank(studentInfo)
    
    VerifyRecordsAreComplete = (blankCount = 0)
End Function

Public Function TrimStringBeforeCharacter(ByVal stringToTrim As String, Optional ByVal trimPoint As String = "(") As String
    Dim charPos As Long
    
    charPos = InStr(stringToTrim, trimPoint)
    If charPos > 0 Then
        stringToTrim = Left$(stringToTrim, charPos - 1)
    End If
    
    TrimStringBeforeCharacter = stringToTrim
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

Public Function IsValueEnglish(ByVal inputText As String) As Boolean
    Dim CharCode As Long
    Dim i As Long
    
    For i = 1 To Len(inputText)
        CharCode = AscW(Mid$(inputText, i, 1))
        
        If CharCode >= &HAC00 And CharCode <= &HD7AF Then
            IsValueEnglish = False
            Exit Function
        End If
    Next i
    
    IsValueEnglish = True
End Function

Public Function GetMergedName(ByVal engName As String, ByVal korName As String) As String
    GetMergedName = engName & " (" & korName & ")"
End Function

Public Function FormatName(ByVal nameValue As String) As String
    Dim msgResult As Long
    
    If Not IsValueEnglish(nameValue) Then
        DisplayMessage "Display.StudentRecords.InvalidEnglishNameEntered", nameValue
        FormatName = vbNullString
        Exit Function
    End If
    
    If ContainsDisallowedChars(nameValue, "English Name") Then
        nameValue = RemoveInvalidCharFromName(nameValue, "English Name")
        If DisplayMessage("Display.StudentRecords.InvalidPunctuationInName", nameValue) = vbNo Then
            FormatName = vbNullString
            Exit Function
        End If
    End If
    
    Select Case Len(nameValue)
        Case 1
            msgResult = DisplayMessage("Display.StudentRecords.EnglishNameIsOneLetter", nameValue)
            If msgResult = vbYes Then
                nameValue = StrConv(nameValue, vbProperCase)
            Else
                nameValue = vbNullString
            End If
        Case 2
            If nameValue = UCase$(nameValue) Then
                msgResult = DisplayMessage("Display.StudentRecords.EnglishNameIsTwoCapitalLetters", nameValue)
                If msgResult = vbNo Then
                    nameValue = StrConv(nameValue, vbProperCase)
                End If
            Else
                msgResult = DisplayMessage("Display.StudentRecords.EnglishNameIsTwoLetters", nameValue)
                If msgResult = vbYes Then
                    nameValue = StrConv(nameValue, vbProperCase)
                Else
                    nameValue = vbNullString
                End If
            End If
        Case Else
            nameValue = StrConv(nameValue, vbProperCase)
    End Select
    
    FormatName = nameValue
End Function

Private Function ContainsDisallowedChars(ByVal valueToSearch As String, ByVal valueType As String) As Boolean
    Const DISALLOWED_NAME_CHARACTERS As String = "[|/\(){}<>'`:;,?~@#$%^&*+=_]"
    
    Select Case valueType
        Case "English Name", "Korean Name"
            ContainsDisallowedChars = (valueToSearch Like "*" & DISALLOWED_NAME_CHARACTERS & "*")
        Case "Comment"
        
    End Select
End Function

Private Function RemoveInvalidCharFromName(ByVal nameValue As String, ByVal valueType As String) As String
    Const DISALLOWED_NAME_CHARACTERS As String = "[|/\(){}<>'`:;,?~@#$%^&*+=_]"
    
    Dim disallowedList As String
    Dim char As String
    Dim i As Long
    
    Select Case valueType
        Case "English Name", "Korean Name"
            disallowedList = DISALLOWED_NAME_CHARACTERS
    End Select
    
    For i = 1 To Len(disallowedList)
        char = Mid$(disallowedList, i, 1)
        nameValue = Replace$(nameValue, char, vbNullString)
    Next i
    
    RemoveInvalidCharFromName = nameValue
End Function

Public Function FormatComment(ByVal commentValue As String) As String
    Const DISALLOWED_FINAL_CHARACTERS As String = "[|/\(){}<>'`:;,?~@#$%^&*+=_-]"
    
    Dim tempString As String
    Dim finalChar As String
    
    tempString = UCase$(Left$(commentValue, 1)) & Mid$(commentValue, 2)
    finalChar = Right$(tempString, 1)
    
    If finalChar <> "." And finalChar <> "!" Then
        Select Case True
            Case finalChar Like "[A-Za-z0-9]"
                tempString = tempString & "."
            Case finalChar Like DISALLOWED_FINAL_CHARACTERS
                tempString = Left$(tempString, Len(tempString) - 1) & "."
                DisplayMessage "Display.StudentRecords.CommentEndsWithInvalidPunctuation", finalChar
            Case (finalChar = "["), (finalChar = "]")
                tempString = Left$(tempString, Len(tempString) - 1) & "."
                DisplayMessage "Display.StudentRecords.CommentEndsWithInvalidPunctuation", finalChar
        End Select
    End If
    
    FormatComment = tempString
End Function

Public Function FormatEvalDate(ByVal dateVal As String) As String
    If IsDate(dateVal) Then
        FormatEvalDate = Format$(CDate(dateVal), GetLocaleDateOrder)
    Else
        If dateVal <> vbNullString Then
            DisplayMessage "Display.StudentRecords.EnterValidDateUponEntry"
        End If
        FormatEvalDate = vbNullString
    End If
End Function

Public Function GetLocaleDateOrder() As String
    Const MMDDYYYY As String = "MMM. DD YYYY"
    Const DDMMYYYY As String = "DD MMM. YYYY"
    Const YYYYMMDD As String = "YYYY MMM. DD"
    
    Dim dateOrder As String
    
    Select Case Application.International(xlDateOrder)
        Case 0
            dateOrder = MMDDYYYY
        Case 1
            dateOrder = DDMMYYYY
        Case 2
            dateOrder = YYYYMMDD
    End Select
    
    GetLocaleDateOrder = dateOrder
End Function

Public Function FormatGrade(ByVal gradeValue As String, Optional ByVal standardizeGradeValues As Boolean = False) As String
    Dim processedGrade As String
    Dim gradeArray As Variant
    
    processedGrade = UCase$(Trim$(gradeValue))
    gradeArray = Array(vbNullString, "C", "B", "B+", "A", "A+")
    
    If processedGrade <> vbNullString Then
        Select Case processedGrade
            Case "A+", "A", "B+", "B", "C"
                FormatGrade = processedGrade
            Case "1" To "5"
                If standardizeGradeValues Then
                    FormatGrade = gradeArray(CLng(processedGrade))
                Else
                    FormatGrade = processedGrade
                End If
            Case Else
                ' If no direct match, attempt to determine intended grade
                 processedGrade = TrimToFinalLetterGrade(processedGrade)
                 
                 If processedGrade = vbNullString Then
                    DisplayMessage "Display.StudentRecords.InvalidScoreUponEntry"
                 End If
                 
                 FormatGrade = processedGrade
        End Select
    End If
End Function

Private Function ValidateClassInfoFromArray(ByVal classInfoData As Variant) As Boolean
    Dim cellValue As String
    Dim labelValue As String
    Dim dataCategory As Long

    ValidateClassInfoFromArray = True ' Assume valid initially

    ' Skip index 2 as it contains no data
    For dataCategory = LBound(classInfoData, 1) To UBound(classInfoData, 1)
        labelValue = Trim$(CStr(classInfoData(dataCategory, 1)))
        cellValue = Trim$(CStr(classInfoData(dataCategory, 3)))

        If cellValue = vbNullString Then
            DisplayMessage "Display.StudentRecords.MissingData", Left$(labelValue, Len(labelValue) - 1)

            If g_UserOptions.EnableLogging Then
                DebugAndLogging GetMsg("Debug.StudentRecords.MissingData", Left$(labelValue, Len(labelValue) - 1))
            End If

            ValidateClassInfoFromArray = False
            Exit Function
        End If

        If dataCategory >= 3 And dataCategory <= 5 Then
            If Not ValidateData(cellValue, labelValue) Then
                If DisplayMessage("Display.StudentRecords.InvalidData", Left$(labelValue, Len(labelValue) - 1)) = vbNo Then
                    If g_UserOptions.EnableLogging Then
                        DebugAndLogging GetMsg("Debug_StudentsRecords_InvalidData", Left$(labelValue, Len(labelValue) - 1))
                    End If
                    
                    ValidateClassInfoFromArray = False
                    Exit Function
                End If
            End If
        End If
    Next dataCategory
End Function

Private Function ValidateStudentInfoFromArray(ByRef ws As Worksheet, ByVal studentData As Variant) As Boolean
    Const HEADER_ROW              As Long = 7
    Const ENGLISH_NAME_COLUMN     As Long = 1
    Const KOREAN_NAME_COLUMN      As Long = 2
    Const GRADE_COLUMNS_START     As Long = 3
    Const GRADE_COLUMNS_END       As Long = 8
    Const COMMENT_COLUMN          As Long = 9
    Const MAX_ENGLISH_NAME_LENGTH As Long = 40
    Const MAX_KOREAN_NAME_LENGTH  As Long = 40
    
    Dim targetCell     As Range
    Dim englishName    As String
    Dim koreanName     As String
    Dim studentName    As String
    Dim categoryName   As String
    Dim currentStudent As Long
    Dim dataCategory   As Long

    ' Assume data is valid by default
    ValidateStudentInfoFromArray = True

    For currentStudent = LBound(studentData, 1) To UBound(studentData, 1)
        englishName = Trim$(CStr(studentData(currentStudent, ENGLISH_NAME_COLUMN)))
        koreanName = Trim$(CStr(studentData(currentStudent, KOREAN_NAME_COLUMN)))
        
        Select Case True
            Case englishName <> vbNullString And koreanName <> vbNullString
                studentName = englishName & "(" & koreanName & ")"
            Case englishName = vbNullString And koreanName = vbNullString
                DisplayMessage "Display.StudentRecords.StudentNameMissing", CStr(currentStudent)
                ValidateStudentInfoFromArray = False
                Exit Function
            Case englishName = vbNullString
                studentName = koreanName & " (Row: " & CStr(currentStudent) & ")"
                DisplayMessage "Display.StudentRecords.StudentNameIncomplete", studentName, "English"
                ValidateStudentInfoFromArray = False
                Exit Function
            Case koreanName = vbNullString
                studentName = englishName & " (Row: " & CStr(currentStudent) & ")"
                DisplayMessage "Display.StudentRecords.StudentNameIncomplete", , studentName, "Korean"
                ValidateStudentInfoFromArray = False
                Exit Function
        End Select
        
        For dataCategory = GRADE_COLUMNS_START To COMMENT_COLUMN
            Set targetCell = ws.Cells.Item(HEADER_ROW + currentStudent, 1 + dataCategory)
            
            categoryName = Trim$(ws.Cells.Item(HEADER_ROW, 1 + dataCategory).Value)

            If IsMissingData(studentData(currentStudent, dataCategory), studentName, categoryName) Then
                ValidateStudentInfoFromArray = False
                Exit Function
            End If

            If Not ValidateData(studentData(currentStudent, dataCategory), categoryName, targetCell) Then
                Select Case dataCategory
                    Case ENGLISH_NAME_COLUMN
                        DisplayMessage "Display.StudentRecords.NameTooLong", studentName, UCase$(categoryName), CStr(MAX_ENGLISH_NAME_LENGTH), Len(studentData(currentStudent, dataCategory)) - MAX_ENGLISH_NAME_LENGTH
                    Case KOREAN_NAME_COLUMN
                        DisplayMessage "Display.StudentRecords.NameTooLong", studentName, UCase$(categoryName), CStr(MAX_KOREAN_NAME_LENGTH), Len(studentData(currentStudent, dataCategory)) - MAX_KOREAN_NAME_LENGTH
                    Case GRADE_COLUMNS_START To GRADE_COLUMNS_END
                        DisplayMessage "Display.StudentRecords.InvalidScore", studentName, UCase$(categoryName)
                    Case COMMENT_COLUMN
                        DisplayMessage "Display.StudentRecords.CommentTooLong", studentName, Len(studentData(currentStudent, dataCategory))
                End Select

                ValidateStudentInfoFromArray = False
                Exit Function
            End If
        Next dataCategory
    Next currentStudent
End Function

Private Function IsMissingData(ByVal dataValue As Variant, ByVal studentName As String, ByVal categoryName As String) As Boolean
    If dataValue = vbNullString Then
        DisplayMessage "Display.StudentRecords.RecordsIncomplete", studentName, UCase$(categoryName)
        IsMissingData = True
    End If
End Function

Private Function ValidateData(ByVal dataValue As String, ByVal dataType As String, Optional ByVal targetCell As Range) As Boolean
    Const MAX_ENGLISH_NAME_LENGTH As Long = 40
    Const MAX_COMMENT_LENGTH      As Long = 960
    Const MIN_SCORE               As Long = 1
    Const MAX_SCORE               As Long = 5
    
    Static validLevels  As Variant
    Static validDays    As Variant
    Static validTimes   As Variant
    Static gradeMapping As Variant
    Static isDeclared   As Boolean

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
            ValidateData = (Len(dataValue) <= MAX_ENGLISH_NAME_LENGTH)
        Case "Level:"
            ValidateData = IsValueValid(validLevels, dataValue)
        Case "Class Days:"
            ValidateData = IsValueValid(validDays, dataValue)
        Case "(Class 1) Time:"
            ValidateData = IsValueValid(validTimes, dataValue)
        Case "Grammar", "Pronunciation", "Fluency", "Manner", "Content", "Overall Effort"
            dataValue = UCase$(dataValue)
            
            If IsValueValid(gradeMapping, dataValue) Then
                WriteNewRangeValue targetCell, dataValue
                ValidateData = True
            ElseIf IsNumeric(dataValue) And _
                   Val(dataValue) >= MIN_SCORE And _
                   Val(dataValue) <= MAX_SCORE Then
                WriteNewRangeValue targetCell, gradeMapping(Val(dataValue) - 1)
                ValidateData = True
            Else
                ValidateData = False
            End If
        Case "Comments (Positive - Negative - Positive)"
            ValidateData = (Len(dataValue) <= MAX_COMMENT_LENGTH)
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

Public Sub WriteNewRangeValue(ByVal valueRange As Range, ByVal newValue As String)
    With valueRange
        If .Value <> newValue Then .Value = newValue
    End With
End Sub