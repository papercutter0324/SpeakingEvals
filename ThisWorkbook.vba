Option Explicit
#Const PRINT_DEBUG_MESSAGES = True
Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
Dim isAppleScriptInstalled As Boolean
Dim isDialogToolkitInstalled As Boolean

Private Sub Workbook_Open()
    Const CURL_COMMAND_TEXT As String = "curl -L -o ~/Library/Application\ Scripts/com.microsoft.Excel/SpeakingEvals.scpt https://github.com/papercutter0324/SpeakingEvals/raw/main/SpeakingEvals.scpt"
    Const CURL_COMMAND_SHAPE As String = "cURL_Command"
    Const WS_INSTRUCTIONS As String = "Instructions"

    Dim ws As Worksheet
    
    Set ws = ActiveWorkbook.Worksheets(WS_INSTRUCTIONS)
    ws.Shapes(CURL_COMMAND_SHAPE).TextFrame2.TextRange.Characters.Text = CURL_COMMAND_TEXT
    SetShapePositions ws
    AutoPopulateEvaluationDateValues
End Sub

Private Sub Workbook_SheetActivate(ByVal ws As Object)
    Const CURL_COMMAND_TEXT As String = "curl -L -o ~/Library/Application\ Scripts/com.microsoft.Excel/SpeakingEvals.scpt https://github.com/papercutter0324/SpeakingEvals/raw/main/SpeakingEvals.scpt"
    Const CURL_COMMAND_SHAPE As String = "cURL_Command"
    Const WS_INSTRUCTIONS As String = "Instructions"
    
    If ws.Name = WS_INSTRUCTIONS Then
        On Error Resume Next
        ws.Shapes(CURL_COMMAND_SHAPE).TextFrame2.TextRange.Characters.Text = CURL_COMMAND_TEXT
        On Error GoTo 0
        SetShapePositions ws
    End If
End Sub

Private Sub SetShapePositions(ByRef ws As Worksheet)
    Dim shp1 As Shape, shp2 As Shape
    Dim shpNamesArray As Variant
    Dim i As Integer
    
    shpNamesArray = Array("Signature_PlaceHolder", "mySignature", "Download Buttons", "MacOS_Command")

    On Error Resume Next
    For i = LBound(shpNamesArray) To UBound(shpNamesArray)
        Set shp1 = ws.Shapes(shpNamesArray(i))
        If Not shp1 Is Nothing Then
            Select Case shp1.Name
                Case "Signature_Placeholder", "mySignature"
                    Set shp2 = ws.Shapes("TS Message")
                    shp1.Top = shp2.Top + ((shp2.Height - shp1.Height) / 2)
                Case "Download Buttons"
                    Set shp2 = ws.Shapes("Seeing the Code")
                    shp1.Top = shp2.Top
                Case "MacOS_Command"
                    Set shp2 = ws.Shapes("MacOS Users")
                    shp1.Top = shp2.Top
            End Select
            shp1.Left = shp2.Left + ((shp2.Width - shp1.Width) / 2)
            Set shp2 = Nothing
        End If
        Set shp1 = Nothing
    Next i
    On Error GoTo 0
End Sub

Private Sub AutoPopulateEvaluationDateValues()
    Dim ws As Worksheet, dateValue As Range
    Dim messageText As String, msgResult As Variant
    Dim i As Long, lastRow As Long
    
    On Error Resume Next
    Application.EnableEvents = False
    For i = 1 To ThisWorkbook.Worksheets.Count
        Set ws = ThisWorkbook.Worksheets(i)
        lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
        If ws.Range("A6").Value = "Evaluation Date:" Then
            SetCorrectDateValidationMessage ws
            
            Set dateValue = ws.Range("C6")
            
            If lastRow < 8 Then
                ' Always refresh the date if not data entered yet
                dateValue.Value = Date
            ElseIf IsEmpty(dateValue) Then
                ' Add a date if student records are present
                dateValue.Value = Date
            ElseIf IsDate(dateValue.Value) And dateValue.Value < Date - 45 Then
                ' Update the date if the current value is over 45 days ago
                dateValue.Value = Date
            ElseIf Not IsDate(dateValue.Value) Then
                ' Display an error message if an invalid date is found
                messageText = "An invalid date has been found on worksheet " & ws.Name & "." & vbNewLine & "Please enter a valid date."
                msgResult = DisplayMessage(messageText, vbInformation, "Invalid Date!")
            End If
        End If
    Next i
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Private Sub SetCorrectDateValidationMessage(ByRef ws As Worksheet)
    Const dayFirstFormat As Date = "01/02/2025"     '02 Jan 2025
    Const monthFirstFormat As Date = "02/01/2025"   '01 Feb 2025
    
    Dim dateFormatStyle As String, dateFormatMessage As String, dateFormula1 As String
    
    ' This can probably be simplified to:
    ' Select Case Application.International(xlDateOrder)
    '    Case 0
    '        dateFormatStyle = "MM/DD/YYYY
    '        dateFormula1 = "1/1/2025"
    '    Case 1
    '        dateFormatStyle = "DD/MM/YYYY
    '        dateFormula1 = ""1/1/2025"
    '    Case 2 ' Check this at home
    '        dateFormatStyle = "YYYY/MM/DD
    '        dateFormula1 = "2025/1/1"
    ' End Select
    dateFormatStyle = IIf(dayFirstFormat < monthFirstFormat, "MM/DD/YYYY", "DD/MM/YYYY")
    dateFormula1 = "1/1/2025"
    dateFormatMessage = "Format:" & vbNewLine & "    " & dateFormatStyle & vbNewLine & vbNewLine & "Only the month and year are displayed, but Excel also needs the day."
    
    On Error Resume Next
    ws.Unprotect
    With ws.Cells(6, 3).Validation
        .Delete ' Clear current validation
        .Add Type:=xlValidateDate, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlGreater, _
             Formula1:=dateFormula1
        .InputTitle = "Evaluation Date"
        .InputMessage = dateFormatMessage
        .ErrorTitle = "Invalid Entry"
        .errorMessage = "Please enter a valid date."
    End With
    ws.Protect
    ws.EnableSelection = xlUnlockedCells
    On Error GoTo 0
End Sub

Sub PrintReports()
    Const REPORT_TEMPLATE As String = "Speaking Evaluation Template.docx"
    Const ERR_INCOMPLETE_RECORDS As String = "incompleteRecords"
    Const ERR_LOADING_WORD As String = "loadingWord"
    Const ERR_LOADING_TEMPLATE As String = "loadingTemplate"
    Const ERR_MISSING_SHAPES As String = "missingTemplateShapes"
    Const MSG_SAVE_FAILED As String = "exportFailed"
    Const MSG_ZIP_FAILED As String = "zipFailed"
    Const MSG_SUCCESS As String = "exportSuccessful"

    Dim ws As Worksheet, wordApp As Object, wordDoc As Object
    Dim templatePath As String, savePath As String, fileName As String
    Dim resultMsg As String, msgToDisplay As String, msgTitle As String, msgType As Integer, msgResult As Variant, dialogSize As Integer
    Dim currentRow As Long, lastRow As Long, firstStudentRecord As Integer
    Dim generateProcess As String, preexistingWordInstance As Boolean, saveResult As Boolean

    Set ws = ActiveSheet
    
    generateProcess = Application.Caller
    If generateProcess = "Button_GenerateReports" Then
        generateProcess = "FinalReports"
    ElseIf generateProcess = "Button_GenerateProofs" Then
        generateProcess = "Proofs"
    Else
        msgToDisplay = "You have clicked an invalid option for creating the reports. This shouldn't be possible unless this file has been altered " & _
                       "in an unintended manner. Please download a new copy of this Excel file, copy over all of the students' records, and try again."
        msgResult = DisplayMessage(msgToDisplay, vbExclamation, "Invalid Selection!")
        Exit Sub
    End If

    #If Mac Then
        RequestInitialFileAndFolderAccess
        isAppleScriptInstalled = CheckForAppleScript()
        If isAppleScriptInstalled Then isDialogToolkitInstalled = CheckForDialogToolkit()
        If isDialogToolkitInstalled Then isDialogToolkitInstalled = CheckForDialogDisplayScript()
    #End If

    If IsTemplateAlreadyOpen(REPORT_TEMPLATE, preexistingWordInstance) Then Exit Sub

    If Not VerifyRecordsAreComplete(ws, lastRow, firstStudentRecord) Then
        resultMsg = ERR_INCOMPLETE_RECORDS
        GoTo Cleanup
    End If

    templatePath = LoadTemplate(REPORT_TEMPLATE)
    If templatePath = "" Then GoTo Cleanup

    savePath = SetSaveLocation(ws, generateProcess)
    If savePath = "" Then GoTo Cleanup

    If Not LoadWord(wordApp, wordDoc, templatePath) Then
        resultMsg = ERR_LOADING_WORD
        GoTo Cleanup
    End If
    
    If wordDoc Is Nothing Then
        resultMsg = ERR_LOADING_TEMPLATE
        GoTo Cleanup
    End If
    
    If Not VerifyAllDocShapesExist(wordDoc) Then
        resultMsg = ERR_MISSING_SHAPES
        GoTo Cleanup
    End If
    
    For currentRow = firstStudentRecord To lastRow
        ClearAllTextBoxes wordDoc
        WriteReport ws, wordApp, wordDoc, generateProcess, currentRow, savePath, fileName, saveResult
    Next currentRow
    
    ws.Activate ' Ensure the right worksheet is being shown when finished.
    
    If Not saveResult Then
        resultMsg = MSG_SAVE_FAILED
        GoTo Cleanup
    End If
    
    #If Mac Then
        If Not isAppleScriptInstalled Then
            resultMsg = MSG_SUCCESS
            GoTo Cleanup
        End If
    #End If
    
    KillWord wordApp, wordDoc, preexistingWordInstance
    resultMsg = MSG_SUCCESS
    
    If generateProcess = "FinalReports" Then
        ZipReports ws, savePath, saveResult
        If Not saveResult Then resultMsg = MSG_ZIP_FAILED
    End If
    
Cleanup:
    Select Case resultMsg
        Case ERR_INCOMPLETE_RECORDS
            msgToDisplay = "One or more fields for missing. Please complete all fields and try again."
            msgTitle = "Missing Data!"
            msgType = vbExclamation
            dialogSize = 230
        Case ERR_LOADING_WORD, ERR_LOADING_TEMPLATE
            msgToDisplay = "There was an error opening MS Word and/or the template. This is sometimes normal MS Office behaviour, so please wait a couple seconds and try again."
            msgTitle = "Error!"
            msgType = vbExclamation
            dialogSize = 360
        Case ERR_MISSING_SHAPES
            msgToDisplay = "There is a error with the template. Please redownload a copy of the original and try again."
            msgTitle = "Error!"
            msgType = vbExclamation
            dialogSize = 210
        Case MSG_SAVE_FAILED
            msgToDisplay = "Export failed. Please ensure all data was entered correctly and try saving to a different folder."
            msgTitle = "Process failed!"
            msgType = vbInformation
            dialogSize = 230
        Case MSG_ZIP_FAILED
            msgToDisplay = "The reports were successfully created, but there was an error when trying to add them into a zip file."
            msgTitle = "Error!"
            msgType = vbInformation
            dialogSize = 270
        Case MSG_SUCCESS
            msgToDisplay = "Export complete!"
            msgTitle = "Process complete!"
            msgType = vbInformation
            dialogSize = 110
    End Select
    
    If dialogSize = 0 Then dialogSize = 250
    If resultMsg <> "" Then msgResult = DisplayMessage(msgToDisplay, msgType, msgTitle, dialogSize)
    KillWord wordApp, wordDoc, preexistingWordInstance
    If isDialogToolkitInstalled Then RemoveDialogToolKit
    Set ws = Nothing
End Sub

Private Function VerifyRecordsAreComplete(ByRef ws As Worksheet, ByRef lastRow As Long, ByRef firstStudentRecord As Integer) As Boolean
    Const CLASS_INFO_FIRST_ROW As Integer = 1
    Const CLASS_INFO_LAST_ROW As Integer = 6
    Const STUDENT_INFO_FIRST_ROW As Integer = 8
    Const STUDENT_INFO_FIRST_COL As Integer = 2
    Const STUDENT_INFO_LAST_COL As Integer = 10
    
    Dim currentRow As Long, currentColumn As Long
    Dim errorMessage As String, msgResult As Variant
    
    ' Set here and passed back to keep declarations organized
    firstStudentRecord = STUDENT_INFO_FIRST_ROW
    
    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.Count, STUDENT_INFO_FIRST_COL).End(xlUp).row
    On Error GoTo 0
    
    If lastRow < STUDENT_INFO_FIRST_ROW Then
        errorMessage = "No students were found!"
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print errorMessage
        #End If
        msgResult = DisplayMessage(errorMessage, vbOKOnly, "Error!", 160)
        VerifyRecordsAreComplete = False
        Exit Function
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Student records found: " & (lastRow - STUDENT_INFO_FIRST_ROW + 1) & vbNewLine & "Beginning validation of entered records."
    #End If
    
    If Not ValidateClassInfo(ws, CLASS_INFO_FIRST_ROW, CLASS_INFO_LAST_ROW) Then
        VerifyRecordsAreComplete = False
        Exit Function
    End If
    
    If Not ValidateStudentInfo(ws, STUDENT_INFO_FIRST_ROW, lastRow, STUDENT_INFO_FIRST_COL, STUDENT_INFO_LAST_COL) Then
        VerifyRecordsAreComplete = False
        Exit Function
    End If

    VerifyRecordsAreComplete = True
End Function

Private Function ValidateData(ByRef currentCell As Range, ByVal dataType As String) As Boolean
    Static validLevels As Variant, validDays As Variant, validTimes As Variant, gradeMapping As Variant, isDeclared As Boolean
    Dim dataValue As String
    
    If Not isDeclared Then
        validLevels = Array("Theseus", "Perseus", "Odysseus", "Hercules", "Artemis", "Hermes", "Apollo", _
                           "Zeus", "E5 Athena", "Helios", "Poseidon", "Gaia", "Hera", "E6 Song's")
        validDays = Array("MonWed", "MonFri", "WedFri", "MWF", "TTh", "MWF (Class 1)", "MWF (Class 2)", _
                         "TTh (Class 1)", "TTh (Class 2)")
        validTimes = Array("4pm", "5pm", "530pm", "6pm", "7pm", "8pm", "830pm", "9pm")
        gradeMapping = Array("C", "B", "B+", "A", "A+")
        isDeclared = True
    End If
    
    Application.EnableEvents = False
    dataValue = Trim(currentCell.Value)
    
    Select Case dataType
        Case "Level:"
            ValidateData = IsValueValid(validLevels, dataValue)
        Case "Class Days:"
            ValidateData = IsValueValid(validDays, dataValue)
        Case "(Class 1) Time:"
            ValidateData = IsValueValid(validTimes, dataValue)
        Case "Grammar", "Pronunciation", "Fluency", "Manner", "Content", "Overall Effort"
            dataValue = UCase(dataValue)
            If IsValueValid(gradeMapping, dataValue) Then
                currentCell.Value = dataValue
                ValidateData = True
            ElseIf IsNumeric(dataValue) And Val(dataValue) >= 1 And Val(dataValue) <= 5 Then
                currentCell.Value = gradeMapping(Val(dataValue) - 1)
                ValidateData = True
            Else
                ValidateData = False
            End If
        Case "Comments"
            ValidateData = (Len(dataValue) < 316)
        Case Else
            ValidateData = False
    End Select
    Application.EnableEvents = True
End Function

Private Function IsValueValid(ByRef dataArray As Variant, ByVal dataValue As String) As Boolean
    Dim i As Integer
    For i = LBound(dataArray) To UBound(dataArray)
        If dataArray(i) = dataValue Then
            IsValueValid = True
            Exit Function
        End If
    Next i
    IsValueValid = False
End Function

Private Function ValidateClassInfo(ByRef ws As Worksheet, ByVal firstRow As Integer, ByVal lastRow As Integer) As Boolean
    Dim currentRow As Integer, errorMessage As String, msgResult As Variant

    For currentRow = firstRow To lastRow
        If IsEmpty(ws.Cells(currentRow, 3).Value) Then
            errorMessage = "Class information incomplete." & vbNewLine & vbNewLine & "Missing: " & Left(ws.Cells(currentRow, 1).Value, Len(ws.Cells(currentRow, 1).Value) - 1)
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print errorMessage
            #End If
            msgResult = DisplayMessage(errorMessage, vbOKOnly, "Error!", 190)
            ValidateClassInfo = False
            Exit Function
        End If

        If currentRow >= 3 And currentRow <= 5 Then
            If Not ValidateData(ws.Cells(currentRow, 3), ws.Cells(currentRow, 1).Value) Then
                errorMessage = "Invalid value for " & Left(ws.Cells(currentRow, 1).Value, Len(ws.Cells(currentRow, 1).Value) - 1) & "." & vbNewLine & vbNewLine & _
                               "Would you like to ignore and continue?"
                If DisplayMessage(errorMessage, vbYesNo, "Error!", 250) = vbNo Then
                    #If PRINT_DEBUG_MESSAGES Then
                        Debug.Print errorMessage
                    #End If
                    ValidateClassInfo = False
                    Exit Function
                End If
            End If
        End If
    Next currentRow

    ValidateClassInfo = True
End Function

Private Function ValidateStudentInfo(ByRef ws As Worksheet, ByVal firstRow As Integer, ByVal lastRow As Integer, ByVal firstCol As Integer, ByVal lastCol As Integer) As Boolean
    Dim currentRow As Integer, currentColumn As Integer
    Dim errorMessage As String, msgResult As Variant, dialogSize As Integer
    
    For currentRow = firstRow To lastRow
        For currentColumn = firstCol To lastCol
            If IsEmpty(ws.Cells(currentRow, currentColumn).Value) Then
                errorMessage = "Student information incomplete." & vbNewLine & vbNewLine & "Missing data for student " & ws.Cells(currentRow, 1).Value & "'s "
                If currentColumn = 2 Then
                    errorMessage = errorMessage & ws.Cells(7, currentColumn).Value & "."
                ElseIf currentColumn >= 4 And currentColumn <= 9 Then
                    errorMessage = errorMessage & "(" & ws.Cells(currentRow, 2).Value & ") " & UCase(ws.Cells(7, currentColumn).Value) & " score."
                ElseIf currentColumn = 10 Then
                    errorMessage = errorMessage & "(" & ws.Cells(currentRow, 2).Value & ") COMMENT."
                End If
                dialogSize = 250
                GoTo ErrorHandler
            End If
            
            If currentColumn >= 4 Then
                If Not ValidateData(ws.Cells(currentRow, currentColumn), ws.Cells(7, currentColumn).Value) Then
                    If currentColumn <> 10 Then
                        errorMessage = "Invalid value entered for student " & ws.Cells(currentRow, 1).Value & "'s (" & ws.Cells(currentRow, 2).Value & ") " & UCase(ws.Cells(7, currentColumn).Value) & " score."
                        dialogSize = 300
                    Else
                        errorMessage = "The COMMENT for student " & ws.Cells(currentRow, 1).Value & "'s (" & ws.Cells(currentRow, 2).Value & ") is too long. Please try to shorten it by " & _
                                        Len(ws.Cells(currentRow, currentColumn).Value) - 315 & " or more characters."
                        dialogSize = 330
                    End If
                    GoTo ErrorHandler
                End If
            End If
        Next currentColumn
    Next currentRow
    
    ValidateStudentInfo = True
    Exit Function
    
ErrorHandler:
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print errorMessage
    #End If
    msgResult = DisplayMessage(errorMessage, vbExclamation, "Error!", dialogSize)
    ValidateStudentInfo = False
End Function

Private Function LoadTemplate(ByVal REPORT_TEMPLATE As String) As String
    Dim templatePath As String, tempTemplatePath As String, destinationPath As String
    Dim msgToDisplay As String, msgResult As Variant
    Dim validTemplateFound As Boolean
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Attempting to load the Speaking Evaluation Template.docx."
    #End If
    
    templatePath = ThisWorkbook.Path & Application.PathSeparator & REPORT_TEMPLATE
    ConvertOneDriveToLocalPath templatePath
    destinationPath = templatePath
    tempTemplatePath = GetTempFilePath(REPORT_TEMPLATE)
    
    DeleteFile tempTemplatePath ' Removing existing file to avoid problems overwriting
    
    If Not IsTemplateValid(templatePath, tempTemplatePath) Then
        msgToDisplay = "No template was found. Process canceled."
        msgResult = DisplayMessage(msgToDisplay, vbOKOnly, "Template Not Found", 150)
        LoadTemplate = ""
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Unable to locate a copy of the template."
        #End If
        Exit Function
    End If
    
    If templatePath = tempTemplatePath Then
        If Not MoveFile(tempTemplatePath, destinationPath) Then
            msgToDisplay = "Failed to move temporary template to final location. Please try downloading the template manually and saving it in this folder."
            msgResult = DisplayMessage(msgToDisplay, vbOKOnly, "Error!", 320)
            LoadTemplate = ""
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "Unable to move the template to the correct location."
            #End If
            Exit Function
        End If
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Template successfully loaded."
    #End If
    
    LoadTemplate = templatePath
End Function

Private Function IsTemplateValid(ByRef templatePath As String, ByVal tempTemplatePath As String) As Boolean
    #If Mac Then
        Dim msgToDisplay As String
        
        If Not isAppleScriptInstalled Then
            IsTemplateValid = (Dir(templatePath) <> "")
            If IsTemplateValid Then
                msgToDisplay = "A template file was found, but its validity cannot be confirmed without SpeakingEvals.scpt. Proceed anyway?"
                IsTemplateValid = (DisplayMessage(msgToDisplay, vbYesNo, "Warning!") = vbYes)
            End If
            Exit Function
        End If
        
        If VerifyTemplateHash(templatePath) Then
            IsTemplateValid = True
            Exit Function
        End If
    #Else
        If Dir(templatePath) <> "" And VerifyTemplateHash(templatePath) Then
            IsTemplateValid = True
            Exit Function
        End If
    #End If
    
    ' Delete invalid and/or non-local copies and grab a fresh copy
    DeleteFile templatePath
    templatePath = tempTemplatePath
    IsTemplateValid = DownloadReportTemplate(templatePath)
End Function

Private Function DownloadReportTemplate(ByVal templatePath As String) As Boolean
    Const REPORT_TEMPLATE_URL As String = "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/Speaking%20Evaluation%20Template.docx"
    Dim downloadResult As Boolean
    
    #If Mac Then
        If isAppleScriptInstalled Then
            On Error Resume Next
            downloadResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DownloadFile", templatePath & APPLE_SCRIPT_SPLIT_KEY & REPORT_TEMPLATE_URL)
            
            If downloadResult Then RequestAdditionalFileAndFolderAccess templatePath
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print IIf(Err.Number = 0, "Download successful.", "Error: " & Err.Description)
            #End If
            On Error GoTo 0
        End If
    #Else
        If CheckForCurl() Then
            downloadResult = DownloadUsingCurl(templatePath, REPORT_TEMPLATE_URL)
        ElseIf CheckForDotNet35() Then
            downloadResult = DownloadUsingDotNet35(templatePath, REPORT_TEMPLATE_URL)
        Else
            downloadResult = False
        End If
    #End If
    
    If downloadResult Then
        DownloadReportTemplate = VerifyTemplateHash(templatePath)
    Else
        DownloadReportTemplate = False
    End If
End Function

Private Function VerifyTemplateHash(ByVal filePath As String) As Boolean
    Const TEMPLATE_HASH As String = "C0343895A881DF739B2B974635A100A6"
    
    #If Mac Then
        Dim msgToDisplay As String, msgResult As Variant
        If Not isAppleScriptInstalled Then
            msgToDisplay = "SpeakingEvals.scpt has not been installed, so the report template's file integrity cannot be validated. The reports will " & _
                           "still be created, but please check that everything looks okay."
            msgResult = DisplayMessage(msgToDisplay, vbOKOnly, "Notice")
            VerifyTemplateHash = True
            Exit Function
        End If
            
        VerifyTemplateHash = AppleScriptTask(APPLE_SCRIPT_FILE, "CompareMD5Hashes", filePath & APPLE_SCRIPT_SPLIT_KEY & TEMPLATE_HASH)
        Exit Function
    #Else
        Dim objShell As Object, shellOutput As String
        
        On Error GoTo ErrorHandler
        Set objShell = CreateObject("WScript.Shell")
        shellOutput = objShell.Exec("cmd /c certutil -hashfile """ & filePath & """ MD5").StdOut.ReadAll
        VerifyTemplateHash = (LCase(TEMPLATE_HASH) = LCase(Trim(Split(shellOutput, vbCrLf)(1))))
    #End If
Cleanup:
    #If Mac Then
    #Else
        Set objShell = Nothing
    #End If
    Exit Function
ErrorHandler:
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Error: " & Err.Number & " - " & Err.Description
    #End If
    VerifyTemplateHash = False
    Resume Cleanup
End Function

Private Function SetSaveLocation(ByRef ws As Object, ByVal saveRoutine As String) As String
    Dim filePath As String

    filePath = ThisWorkbook.Path & Application.PathSeparator & GenerateSaveFolderName(ws) & Application.PathSeparator
    ConvertOneDriveToLocalPath filePath

    ' Check if folder already exists. If yes, delete for user simplicity
    If DoesFolderExist(filePath) Then DeleteExistingFolder filePath
    CreateSaveFolder filePath
    
    If saveRoutine = "Proofs" Then
        filePath = filePath & "Proofs"
        If DoesFolderExist(filePath) Then DeleteExistingFolder filePath
        CreateSaveFolder filePath
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Saving reports in: " & vbNewLine & "    " & filePath
    #End If
    
    SetSaveLocation = filePath
End Function

Private Function GenerateSaveFolderName(ByRef ws As Worksheet) As String
    Dim classIdentifier As String
    
    Select Case ws.Cells(4, 3).Value
        Case "MonWed"
            classIdentifier = "MW - " & ws.Cells(5, 3).Value
        Case "MonFri"
            classIdentifier = "MF - " & ws.Cells(5, 3).Value
        Case "WedFri"
            classIdentifier = "WF - " & ws.Cells(5, 3).Value
        Case "MWF"
            classIdentifier = "MWF - " & ws.Cells(5, 3).Value
        Case "TTh"
            classIdentifier = "TTh - " & ws.Cells(5, 3).Value
        Case "MWF (Class 1)": classIdentifier = "MWF-1"
        Case "MWF (Class 2)": classIdentifier = "MWF-2"
        Case "TTh (Class 1)": classIdentifier = "TTh-1"
        Case "TTh (Class 2)": classIdentifier = "TTh-2"
    End Select
    
    GenerateSaveFolderName = ws.Cells(3, 3).Value & " (" & classIdentifier & ")"
End Function

Private Function DoesFolderExist(ByVal filePath As String) As Boolean
    #If Mac Then
        If isAppleScriptInstalled Then
            DoesFolderExist = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFolderExist", filePath)
            Exit Function
        End If
    #End If

    DoesFolderExist = (Dir(filePath, vbDirectory) <> "")
End Function

Private Sub CreateSaveFolder(ByRef filePath As String)
    If Right(filePath, 1) = Application.PathSeparator Then
        filePath = Left(filePath, Len(filePath) - 1)
    End If

    On Error Resume Next
    #If Mac Then
        Dim msgToDisplay As String, msgTitle As String, msgResult As Variant
        Dim scriptResult As Boolean

        If isAppleScriptInstalled Then
            scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CreateFolder", filePath)
        Else
            If Dir(filePath, vbDirectory) = "" Then MkDir filePath
            If Dir(filePath & "/*") <> "" Then
                msgToDisplay = "It appears some files still exist in " & filePath & ". " & vbNewLine & vbNewLine & "The new reports will be generated, but " & _
                               "the old files will not be deleted and may be overwritten."
                msgResult = DisplayMessage(msgToDisplay, vbOKOnly, "Notice")
            End If
        End If

        RequestAdditionalFileAndFolderAccess filePath & Application.PathSeparator
    #Else
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder filePath
        Set fso = Nothing
    #End If
    On Error GoTo 0

    If Right(filePath, 1) <> Application.PathSeparator Then
        filePath = filePath & Application.PathSeparator
    End If
End Sub

Private Function IsTemplateAlreadyOpen(ByVal REPORT_TEMPLATE As String, ByRef preexistingWordInstance As Boolean) As Boolean
    Dim wordApp As Object, wordDoc As Object
    Dim templatePath As String, templateIsOpen As Boolean
    Dim warningMsg As String
    
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    Err.Clear
    
    If Not wordApp Is Nothing Then
        preexistingWordInstance = True
    
        If Left(ThisWorkbook.Path, 23) = "https://d.docs.live.net" Or Left(ThisWorkbook.Path, 11) = "OneDrive://" Then
            templatePath = ThisWorkbook.Path & "/" & REPORT_TEMPLATE
        Else
            templatePath = ThisWorkbook.Path & Application.PathSeparator & REPORT_TEMPLATE
        End If
        
        For Each wordDoc In wordApp.Documents
            If StrComp(wordDoc.FullName, templatePath, vbTextCompare) = 0 Then
                templateIsOpen = True
                Exit For
            End If
        Next wordDoc
    End If
    
    If templateIsOpen Then
        warningMsg = "An open instance of MS Word has been detected. Please save any open files before continuing." & vbNewLine & vbNewLine & _
                     "Click OK to automatically close Word and continue, or click Cancel to finish and save your work."
        If DisplayMessage(warningMsg, vbCritical, "Error Loading Word", 310) = vbOK Then
            wordDoc.Close SaveChanges:=False
            templateIsOpen = False
        End If
    End If
    On Error GoTo 0
    
    Set wordDoc = Nothing
    Set wordApp = Nothing
    IsTemplateAlreadyOpen = templateIsOpen
End Function

Private Function LoadWord(ByRef wordApp As Object, ByRef wordDoc As Object, ByVal templatePath As String) As Boolean
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    Err.Clear
    On Error GoTo ErrorHandler
    
    ' Open a new instance of Word if needed
    #If Mac Then
        Dim appleScriptResult As String, errorMsg As String, msgResult As Variant
        
        If isAppleScriptInstalled Then
            appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "LoadApplication", "Microsoft Word")
            
            #If PRINT_DEBUG_MESSAGES Then
                If appleScriptResult <> "" Then Debug.Print appleScriptResult
            #End If
            
            appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "IsAppLoaded", "Microsoft Word")
            
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print appleScriptResult
            #End If
            
            Set wordApp = GetObject(, "Word.Application")
        End If
    #End If
    If wordApp Is Nothing Then Set wordApp = CreateObject("Word.Application")
    
    ' Make the process visible so users understand their computer isn't frozen
    wordApp.Visible = True
    wordApp.ScreenUpdating = True
    
    If Not wordApp Is Nothing Then Set wordDoc = wordApp.Documents.Open(templatePath)
    LoadWord = (Not wordApp Is Nothing)
    Exit Function
ErrorHandler:
    #If Mac Then
        errorMsg = "An error occurred while trying to load Microsoft Word. This is usually a result of a quirk in MacOS. Try creating the reports again, and it should work fine." & vbNewLine & vbNewLine & _
        "If the problem persists, please take a picture of the following error message and ask your team leader to send it to Warren at Bundang." & vbNewLine & vbNewLine & _
        "VBA Error " & Err.Number & ": " & Err.Description
        
        If isAppleScriptInstalled Then errorMsg = errorMsg & vbNewLine & "AppleScript Error: " & appleScriptResult
        msgResult = DisplayMessage(errorMsg, vbOKOnly, "Error Loading Word", 470)
    #End If
    LoadWord = False
End Function

Private Function VerifyAllDocShapesExist(ByRef wordDoc As Object) As Boolean
    Dim shp As Shape, shapeNames As Variant
    Dim msgToDisplay As String, msgResult As Variant
    Dim i As Integer
    
    shapeNames = Array("English_Name", "Korean_Name", "Grade", "Native_Teacher", "Korean_Teacher", "Date", _
                       "Grammar_A+", "Grammar_A", "Grammar_B+", "Grammar_B", "Grammar_C", _
                       "Pronunciation_A+", "Pronunciation_A", "Pronunciation_B+", "Pronunciation_B", "Pronunciation_C", _
                       "Fluency_A+", "Fluency_A", "Fluency_B+", "Fluency_B", "Fluency_C", _
                       "Manner_A+", "Manner_A", "Manner_B+", "Manner_B", "Manner_C", _
                       "Content_A+", "Content_A", "Content_B+", "Content_B", "Content_C", _
                       "Effort_A+", "Effort_A", "Effort_B+", "Effort_B", "Effort_C", _
                       "Comments", "Overall_Grade")
                       
    For i = LBound(shapeNames) To UBound(shapeNames)
        If Not WordDocShapeExists(wordDoc, shapeNames(i)) Then
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "Missing shape: " & shapeNames(i)
            #End If
            
            msgToDisplay = "There is a critical error with the template. Please redownload a copy of the original and try again."
            msgResult = DisplayMessage(msgToDisplay, vbOKOnly, "Error!", 300)
            VerifyAllDocShapesExist = False
            Exit Function
        End If
    Next i
                       
    VerifyAllDocShapesExist = True
End Function

Private Function WordDocShapeExists(ByRef wordDoc As Object, ByVal shapeName As String) As Boolean
    Dim shp As Object, grpItem As Object
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Search for shape: " & shapeName
    #End If
    
    On Error Resume Next
    For Each shp In wordDoc.Shapes
        If shp.Type = msoGroup Then
            For Each grpItem In shp.GroupItems
                If grpItem.Name = shapeName Then
                    WordDocShapeExists = True
                    Exit Function
                End If
            Next grpItem
        End If
    Next shp
    On Error GoTo 0
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Unable to find shape: " & shapeName
    #End If
    
    WordDocShapeExists = False
End Function

Private Sub ClearAllTextBoxes(wordDoc As Object)
    Dim shp As Object, grpItem As Object
    
    For Each shp In wordDoc.Shapes
        If shp.Type = msoGroup Then
            For Each grpItem In shp.GroupItems
                If grpItem.Type = msoTextBox Or grpItem.Type = msoAutoShape Then
                    grpItem.TextFrame.TextRange.Text = ""
                End If
            Next grpItem
        End If
    Next shp
End Sub

Private Sub WriteReport(ByRef ws As Object, ByRef wordApp As Object, ByRef wordDoc As Object, ByVal generateProcess As String, ByVal currentRow As Integer, ByVal savePath As String, ByRef fileName As String, ByRef saveResult As Boolean)
    Dim nativeTeacher As String, koreanTeacher As String, classLevel As String, classTime As String, evalDate As String
    Dim englishName As String, koreanName As String, grammarScore As String, pronunciationScore As String, fluencyScore As String
    Dim mannerScore As String, contentScore As String, effortScore As String, commentText As String, overallGrade As String
    
    ' Data applicable to all reports
    nativeTeacher = ws.Cells(1, 3).Value
    koreanTeacher = ws.Cells(2, 3).Value
    classLevel = ws.Cells(3, 3).Value
    classTime = ws.Cells(4, 3).Value & "-" & ws.Cells(5, 3).Value
    evalDate = Format(ws.Cells(6, 3).Value, "MMM. YYYY")
    
    ' Data specific to each student
    englishName = ws.Cells(currentRow, 2).Value
    koreanName = ws.Cells(currentRow, 3).Value
    grammarScore = ws.Cells(currentRow, 4).Value
    pronunciationScore = ws.Cells(currentRow, 5).Value
    fluencyScore = ws.Cells(currentRow, 6).Value
    mannerScore = ws.Cells(currentRow, 7).Value
    contentScore = ws.Cells(currentRow, 8).Value
    effortScore = ws.Cells(currentRow, 9).Value
    commentText = ws.Cells(currentRow, 10).Value
    overallGrade = CalculateOverallGrade(ws, currentRow)
    
    fileName = koreanTeacher & "(" & classTime & ")" & " - " & koreanName & "(" & englishName & ")"
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Saving to: " & savePath & vbNewLine & "Saving as: " & fileName
    #End If
    
    With wordDoc
        .Shapes("Report_Header").GroupItems("English_Name").TextFrame.TextRange.Text = englishName
        .Shapes("Report_Header").GroupItems("Korean_Name").TextFrame.TextRange.Text = koreanName
        .Shapes("Report_Header").GroupItems("Grade").TextFrame.TextRange.Text = classLevel
        .Shapes("Report_Header").GroupItems("Native_Teacher").TextFrame.TextRange.Text = nativeTeacher
        .Shapes("Report_Header").GroupItems("Korean_Teacher").TextFrame.TextRange.Text = koreanTeacher
        .Shapes("Report_Header").GroupItems("Date").TextFrame.TextRange.Text = evalDate
        .Shapes("Grammar_Scores").GroupItems("Grammar_" & grammarScore).TextFrame.TextRange.Text = grammarScore
        .Shapes("Pronunciation_Scores").GroupItems("Pronunciation_" & pronunciationScore).TextFrame.TextRange.Text = pronunciationScore
        .Shapes("Fluency_Scores").GroupItems("Fluency_" & fluencyScore).TextFrame.TextRange.Text = fluencyScore
        .Shapes("Manner_Scores").GroupItems("Manner_" & mannerScore).TextFrame.TextRange.Text = mannerScore
        .Shapes("Content_Scores").GroupItems("Content_" & contentScore).TextFrame.TextRange.Text = contentScore
        .Shapes("Effort_Scores").GroupItems("Effort_" & effortScore).TextFrame.TextRange.Text = effortScore
        .Shapes("Report_Footer").GroupItems("Comments").TextFrame.TextRange.Text = commentText
        .Shapes("Report_Footer").GroupItems("Overall_Grade").TextFrame.TextRange.Text = overallGrade
    End With
    
    On Error Resume Next
    If wordDoc.Shapes("Signature") Is Nothing Then InsertSignature wordDoc
    saveResult = SaveToFile(wordDoc, generateProcess, savePath, fileName)
End Sub

Private Function SaveToFile(ByRef wordDoc As Object, ByVal saveRoutine As String, ByVal savePath As String, ByVal fileName As String) As Boolean
    On Error Resume Next
    If saveRoutine = "Proofs" Then
        wordDoc.SaveAs2 fileName:=(savePath & fileName & ".docx"), FileFormat:=16, AddtoRecentFiles:=False, EmbedTrueTypeFonts:=True
    Else
        #If Mac Then
            ' Export to PDF is a bit flaky on MacOS, so we need to do a full SaveAs2. Only results in a minimal time loss.
            wordDoc.SaveAs2 fileName:=(savePath & fileName & ".pdf"), FileFormat:=17, AddtoRecentFiles:=False, EmbedTrueTypeFonts:=True
        #Else
            wordDoc.ExportAsFixedFormat OutputFileName:=(savePath & fileName & ".pdf"), ExportFormat:=17, BitmapMissingFonts:=True
        #End If
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        If Err.Number = 0 Then
            Debug.Print "Successfully saved: " & fileName
        Else
            Debug.Print "Failed to save." & "Error Number: " & Err.Number & vbNewLine & "Error Description: " & Err.Description
        End If
    #End If
    
    SaveToFile = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function CalculateOverallGrade(ByRef ws As Worksheet, ByVal currentRow As Integer) As String
    Dim scoreRange As Range, gradeCell As Range
    Dim totalScore As Double, avgScore As Double
    Dim roundedScore As Integer, numericScore As Integer
    
    Set scoreRange = ws.Range("D" & currentRow & ":I" & currentRow)
    totalScore = 0
    
    For Each gradeCell In scoreRange
        Select Case gradeCell.Value
            Case "A+": numericScore = 5
            Case "A": numericScore = 4
            Case "B+": numericScore = 3
            Case "B": numericScore = 2
            Case "C": numericScore = 1
        End Select
        totalScore = totalScore + numericScore
    Next gradeCell
    
    ' Be a little generous with the score rounding. They're young afterall.
    avgScore = totalScore / 6
    If avgScore - Int(avgScore) >= 0.4 Then
        roundedScore = Int(avgScore) + 1
    Else
        roundedScore = Int(avgScore)
    End If
    
    Select Case roundedScore
        Case 5: CalculateOverallGrade = "A+"
        Case 4: CalculateOverallGrade = "A"
        Case 3: CalculateOverallGrade = "B+"
        Case 2: CalculateOverallGrade = "B"
        Case 1: CalculateOverallGrade = "C"
    End Select
End Function

Private Sub InsertSignature(ByRef wordDoc As Object)
    Dim shp As Shape, newImageShape As Object
    Dim newImageWidth As Double, newImageHeight As Double
    Dim imageWidth As Double, imageHeight As Double, aspectRatio As Double
    Dim signatureFound As Boolean
    
    Const SIGNATURE_SHAPE_NAME As String = "mySignature"
    Const SIGNATURE_PNG_FILENAME As String = "mySignature.png"
    Const SIGNATURE_JPG_FILENAME As String = "mySignature.jpg"
    
    ' These numbers make no sense, but they work.
    Const ABSOLUTE_LEFT As Double = 332.4
    Const ABSOLUTE_TOP As Double = 684
    Const MAX_WIDTH As Double = 144
    Const MAX_HEIGHT As Double = 40
    
    Static signaturePath As String
    Static newImagePath As String
    Static useEmbeddedSignature As Boolean
    
    If signaturePath = "" Then
        signaturePath = ThisWorkbook.Path & Application.PathSeparator
        ConvertOneDriveToLocalPath signaturePath
    End If
    
    On Error Resume Next
    useEmbeddedSignature = (Not ThisWorkbook.Sheets("Instructions").Shapes("mySignature") Is Nothing)
    On Error GoTo 0
     
    If newImagePath = "" Then
        If useEmbeddedSignature Then
            SaveSignature SIGNATURE_SHAPE_NAME, newImagePath
        ElseIf isAppleScriptInstalled Then
            #If Mac Then
                newImagePath = AppleScriptTask(APPLE_SCRIPT_FILE, "FindSignature", signaturePath)
                If newImagePath = "" Then Exit Sub
                signatureFound = True
            #End If
        ElseIf Not signatureFound Then
            If Dir(signaturePath & SIGNATURE_PNG_FILENAME) <> "" Then
                newImagePath = signaturePath & SIGNATURE_PNG_FILENAME
            ElseIf Dir(signaturePath & SIGNATURE_JPG_FILENAME) <> "" Then
                newImagePath = signaturePath & SIGNATURE_JPG_FILENAME
            Else
                Exit Sub
            End If
        End If
    End If
    
    Set newImageShape = wordDoc.Shapes.AddPicture(fileName:=newImagePath, LinkToFile:=False, SaveWithDocument:=True)
    newImageShape.Name = "Signature"
    
    ' Maintain the aspect ratio and resize if needed
    aspectRatio = newImageShape.Width / newImageShape.Height
    If MAX_WIDTH / MAX_HEIGHT > aspectRatio Then
        ' Adjust width to fit within max height
        imageWidth = MAX_HEIGHT * aspectRatio
        imageHeight = MAX_HEIGHT
    Else
        ' Adjust height to fit within max width
        imageWidth = MAX_WIDTH
        imageHeight = MAX_WIDTH / aspectRatio
    End If

    ' Position and resize the image
    With newImageShape
        .LockAspectRatio = msoTrue
        .Left = ABSOLUTE_LEFT
        .Top = ABSOLUTE_TOP
        .Width = imageWidth
        .RelativeHorizontalPosition = 1
        .RelativeVerticalPosition = 1
    End With
End Sub

Private Sub SaveSignature(ByVal SIGNATURE_SHAPE_NAME As String, ByRef savePath As String)
    Dim signSheet As Worksheet, tempSheet As Worksheet
    Dim signatureshp As Shape, chrt As ChartObject
    
    Sheets.Add(, Sheets(Sheets.Count)).Name = "Temp_signature"
    Set tempSheet = Sheets("Temp_signature")
    tempSheet.Select
    
    Set signatureshp = ThisWorkbook.Worksheets("Instructions").Shapes(SIGNATURE_SHAPE_NAME)
    signatureshp.Copy
    
    savePath = GetTempFilePath("tempSignature.png")
    ConvertOneDriveToLocalPath savePath
    
    On Error Resume Next
    Application.DisplayAlerts = False
    With tempSheet.ChartObjects.Add(Left:=tempSheet.Range("B2").Left, Top:=tempSheet.Range("B2").Top, _
                                    Width:=signatureshp.Width, Height:=signatureshp.Height)
        .Activate
        .Chart.Paste
        .Chart.ChartArea.Format.Line.Visible = msoFalse
        .Chart.Export savePath, "png"
        .Delete
    End With
    tempSheet.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Sub ZipReports(ByRef ws As Worksheet, ByVal savePath As String, ByRef saveResult As Boolean)
    Dim zipPath As Variant, zipName As Variant, pdfPath As Variant
    Dim errDescription As String
    
    On Error Resume Next
    If Right(savePath, 1) <> Application.PathSeparator Then savePath = savePath & Application.PathSeparator
    
    zipName = ws.Cells(3, 3).Value & " (" & ws.Cells(2, 3).Value & " " & ws.Cells(4, 3).Value & ").zip"
    zipPath = savePath & zipName
    
    #If Mac Then
        Dim scriptResult As String
        
        If Not isAppleScriptInstalled Then
            saveResult = False
            Exit Sub
        End If
        
        scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CreateZipFile", savePath & APPLE_SCRIPT_SPLIT_KEY & zipPath)
        
        If scriptResult <> "Success" Then
            errDescription = scriptResult
            saveResult = False
        Else
            saveResult = True
        End If
    #Else
        Dim shellApp As Object
        
        ' Create an empty ZIP file
        If Len(Dir(zipPath)) > 0 Then Kill zipPath
        Open zipPath For Output As #1
        Print #1, "PK" & Chr(5) & Chr(6) & String(18, vbNullChar)
        Close #1
        
        Set shellApp = CreateObject("Shell.Application")
        pdfPath = Dir(savePath & "*.pdf") ' Only target PDF files
        
        Do While pdfPath <> ""
            shellApp.Namespace(zipPath).CopyHere savePath & pdfPath
            Application.Wait Now + TimeValue("0:00:01") ' Delay to allow compression
            pdfPath = Dir ' Get the next PDF file
        Loop
        
        Set shellApp = Nothing
        If Err.Number <> 0 Then errDescription = Err.Description
        saveResult = (Err.Number = 0)
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        If saveResult Then
            Debug.Print "Zip file successfully created."
        Else
            Debug.Print "There was an error creating the Zip file." & vbNewLine & "Error: " & errDescription
        End If
    #End If
    On Error GoTo 0
End Sub

Private Sub KillWord(ByRef wordApp As Object, ByRef wordDoc As Object, ByVal preexistingWordInstance As Boolean)
    On Error Resume Next
    If Not wordDoc Is Nothing Then
        wordDoc.Close SaveChanges:=False
        Set wordDoc = Nothing
    End If
    
    If Not preexistingWordInstance Then
        If Not wordApp Is Nothing Then wordApp.Quit
        Set wordApp = Nothing
    
        #If Mac Then
            Dim closeResult As String
            
            If isAppleScriptInstalled Then
                closeResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CloseWord", closeResult)
                #If PRINT_DEBUG_MESSAGES Then
                    Debug.Print closeResult
                #End If
            End If
        #End If
    End If
    On Error GoTo 0
End Sub

#If Mac Then
Private Function CheckForAppleScript() As Boolean
    Dim appleScriptPath As String, appleScriptResult As Boolean
    
    appleScriptPath = "/Users/" & Environ("USER") & "/Library/Application Scripts/com.microsoft.Excel/" & APPLE_SCRIPT_FILE
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Locating " & APPLE_SCRIPT_FILE & vbNewLine & "Searching: " & appleScriptPath
    #End If
    
    On Error Resume Next
    appleScriptResult = (Dir(appleScriptPath, vbDirectory) = APPLE_SCRIPT_FILE)
    On Error GoTo 0
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Found: " & appleScriptResult
    #End If
    
    If appleScriptResult Then
        CheckForAppleScriptUpdate
    End If
    
    CheckForAppleScript = appleScriptResult
End Function

Private Sub CheckForAppleScriptUpdate()
    Dim scriptFolder As String, destinationPath As String
    Dim currentScriptVersion As Long, downloadedScriptVersion As Long
    Dim appleScriptResult As Boolean
    
    Const APPLE_SCRIPT_URL As String = "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/SpeakingEvals.scpt"
    Const TEMP_APPLE_SCRIPT As String = "SpeakingEvals-Temp.scpt"
    Const OLD_APPLE_SCRIPT As String = "SpeakingEvals-Old.scpt"
    
    scriptFolder = "/Users/" & Environ("USER") & "/Library/Application Scripts/com.microsoft.Excel/"
    destinationPath = scriptFolder & TEMP_APPLE_SCRIPT
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Checking if an update is available."
    #End If
    
    On Error GoTo ErrorHandler
    
    appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DownloadFile", destinationPath & APPLE_SCRIPT_SPLIT_KEY & APPLE_SCRIPT_URL)
    If Not appleScriptResult Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Unable to download new " & APPLE_SCRIPT_FILE
        #End If
        GoTo ErrorHandler
    End If
    
    currentScriptVersion = AppleScriptTask(APPLE_SCRIPT_FILE, "GetScriptVersionNumber", "")
    downloadedScriptVersion = AppleScriptTask(TEMP_APPLE_SCRIPT, "GetScriptVersionNumber", "")
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Installed Version: " & currentScriptVersion
        Debug.Print "Latest Version: " & downloadedScriptVersion
    #End If
    
    If downloadedScriptVersion <= currentScriptVersion Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Installed version is up-to-date."
        #End If
        GoTo Cleanup
    End If
    
    appleScriptResult = AppleScriptTask(TEMP_APPLE_SCRIPT, "RenameFile", scriptFolder & APPLE_SCRIPT_FILE & APPLE_SCRIPT_SPLIT_KEY & scriptFolder & OLD_APPLE_SCRIPT)
    If appleScriptResult Then appleScriptResult = AppleScriptTask(OLD_APPLE_SCRIPT, "RenameFile", scriptFolder & TEMP_APPLE_SCRIPT & APPLE_SCRIPT_SPLIT_KEY & scriptFolder & APPLE_SCRIPT_FILE)
    If appleScriptResult Then appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", scriptFolder & OLD_APPLE_SCRIPT)
    If Not appleScriptResult Then GoTo ErrorHandler
    
    #If PRINT_DEBUG_MESSAGES Then
        If appleScriptResult Then Debug.Print "Update complete."
    #End If
    
Cleanup:
    #If PRINT_DEBUG_MESSAGES Then
        If appleScriptResult Then Debug.Print "Beginning clean up process."
    #End If
    
    On Error Resume Next
    appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", scriptFolder & TEMP_APPLE_SCRIPT)
    If appleScriptResult Then appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", scriptFolder & TEMP_APPLE_SCRIPT)
    
    appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", scriptFolder & OLD_APPLE_SCRIPT)
    If appleScriptResult Then appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", scriptFolder & OLD_APPLE_SCRIPT)
    On Error GoTo 0
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Finished clean up."
    #End If
    Exit Sub
    
ErrorHandler:
    #If PRINT_DEBUG_MESSAGES Then
        If Err.Number <> 0 Then Debug.Print "Error during the update process."
        If Err.Description <> "" Then Debug.Print "Error: " & Err.Description
    #End If
    Resume Cleanup
End Sub

Private Function CheckForDialogToolkit() As Boolean
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Checking for presence of Dialog Toolkit Plus."
    #End If
    
    CheckForDialogToolkit = AppleScriptTask(APPLE_SCRIPT_FILE, "InstallDialogToolkitPlus", "paramString")
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Status: " & CheckForDialogToolkit
    #End If
End Function

Private Function CheckForDialogDisplayScript() As Boolean
    Dim resourcesFolder As String
    Dim dialogScriptIsInstalled As Boolean
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Checking for presence of DialogDisplay.scpt."
    #End If
    
    ' Needed to check for a local copy in case no internet connection available
    resourcesFolder = ThisWorkbook.Path & "/Resources"
    ConvertOneDriveToLocalPath resourcesFolder
    
    dialogScriptIsInstalled = AppleScriptTask(APPLE_SCRIPT_FILE, "InstallDialogDisplayScript", resourcesFolder)
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Status: " & dialogScriptIsInstalled
    #End If
    
    CheckForDialogDisplayScript = dialogScriptIsInstalled
End Function

Private Sub RemoveDialogToolKit()
    Dim dialogToolktPlusScript As String, resourcesFolder As String, scriptResult As Boolean
        
    dialogToolktPlusScript = "/Users/" & Environ("USER") & "/Library/Script Libraries/Dialog Toolkit Plus.scptd"
    resourcesFolder = ThisWorkbook.Path & "/Resources"
    ConvertOneDriveToLocalPath resourcesFolder
    
    scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFolderExist", resourcesFolder)
    If Not scriptResult Then scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CreateFolder", resourcesFolder)
    If scriptResult Then scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CopyFolder", dialogToolktPlusScript & APPLE_SCRIPT_SPLIT_KEY & resourcesFolder & "/Dialog Toolkit Plus.scptd")
    If scriptResult Then scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFolder", dialogToolktPlusScript)
End Sub

Private Sub RequestInitialFileAndFolderAccess()
    Dim workingFolder As String, tempFolder As String
    Dim filePermissionCandidates As Variant
    Dim fileAccessGranted As Boolean
    
    workingFolder = ThisWorkbook.Path
    tempFolder = Environ("TMPDIR")
    
    ConvertOneDriveToLocalPath workingFolder
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Requesting access to: " & vbNewLine & _
                    "    " & workingFolder & vbNewLine & _
                    "    " & tempFolder
    #End If
    
    filePermissionCandidates = Array(workingFolder, tempFolder)
    fileAccessGranted = GrantAccessToMultipleFiles(filePermissionCandidates)
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Access granted: " & fileAccessGranted
    #End If
End Sub

Private Sub RequestAdditionalFileAndFolderAccess(ByVal newPath As String)
    Dim filePermissionCandidates As Variant
    Dim fileAccessGranted As Boolean
     
    ConvertOneDriveToLocalPath newPath
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Requesting file & folder access to: " & newPath
    #End If
    
    filePermissionCandidates = Array(newPath)
    fileAccessGranted = GrantAccessToMultipleFiles(filePermissionCandidates)
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Accress granted: " & fileAccessGranted
    #End If
End Sub
#Else
Private Function CheckForCurl() As Boolean
    Dim objShell As Object, objExec As Object
    Dim checkResult As Boolean
    Dim output As String
    
    On Error GoTo ErrorHandler
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Checking if curl.exe is installed."
    #End If
    
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec("cmd /c curl.exe --version")
    
    If Not objExec Is Nothing Then
        Do While Not objExec.StdOut.AtEndOfStream
            output = output & objExec.StdOut.ReadLine() & vbNewLine
        Loop
        checkResult = ((InStr(output, "curl")) > 0)
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print IIf(checkResult, "   Installed.", "   Not installed. Falling back to .Net.")
    #End If
    
    CheckForCurl = checkResult
Cleanup:
    If Not objExec Is Nothing Then Set objExec = Nothing
    If Not objShell Is Nothing Then Set objShell = Nothing
    Exit Function
ErrorHandler:
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Error while checking for curl.exe: " & Err.Description
    #End If
    CheckForCurl = False
    Resume Cleanup
End Function

Private Function DownloadUsingCurl(ByVal templatePath As String, ByVal REPORT_TEMPLATE_URL As String) As Boolean
    Dim objShell As Object, fso As Object
    Dim downloadCommand As String
    
    On Error Resume Next
    Set objShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    downloadCommand = "cmd /c curl.exe -o """ & templatePath & """ """ & REPORT_TEMPLATE_URL & """"
    objShell.Run downloadCommand, 0, True
    DownloadUsingCurl = fso.FileExists(templatePath)
    
    #If PRINT_DEBUG_MESSAGES Then
        If Not DownloadUsingCurl Then Debug.Print "curl download failed for " & REPORT_TEMPLATE_URL
    #End If
    
    Set objShell = Nothing
    Set fso = Nothing
    On Error GoTo 0
End Function

Private Function CheckForDotNet35() As Boolean
    Dim frameworkPath As String
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Verifying that Microsoft DotNet 3.5 is installed."
    #End If
    
    On Error GoTo ErrorHandler
    frameworkPath = Environ$("systemroot") & "\Microsoft.NET\Framework\v3.5"
    CheckForDotNet35 = Dir$(frameworkPath, vbDirectory) <> vbNullString
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "   Checking path: " & frameworkPath & vbNewLine & _
                    "   Installed: " & CheckForDotNet35
    #End If
    
    Exit Function
ErrorHandler:
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Error while checking for .NET 3.5: " & Err.Description
    #End If
    CheckForDotNet35 = False
End Function

Private Function DownloadUsingDotNet35(ByVal templatePath As String, ByVal REPORT_TEMPLATE_URL As String) As Boolean
    Dim xmlHTTP As Object, fileStream As Object
    
    On Error Resume Next
    Set xmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Set fileStream = CreateObject("ADODB.Stream")
    
    xmlHTTP.Open "Get", REPORT_TEMPLATE_URL, False
    xmlHTTP.Send
    
    If xmlHTTP.Status = 200 Then
        fileStream.Open
        fileStream.Type = 1 ' Binary
        fileStream.Write xmlHTTP.responseBody
        fileStream.SaveToFile templatePath, 2 ' Overwrite existing, if somehow present
        fileStream.Close
        DownloadUsingDotNet35 = True
    Else
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "HTTP request failed. Status: " & xmlHTTP.Status & " - " & xmlHTTP.StatusText
        #End If
        DownloadUsingDotNet35 = False
    End If
    
    Set xmlHTTP = Nothing
    Set fileStream = Nothing
    On Error GoTo 0
End Function
#End If

Private Sub ConvertOneDriveToLocalPath(ByRef selectedPath As String)
    Dim i As Integer
    
    ' Cloud storage apps like OneDrive sometimes complicate where/how files are saved. Below is a reference
    ' to track and help add support foradditionalcloud storage providers.
    ' OneDrive
        ' Local Paths:      "/Users/" & Environ("USER") & "/Library/CloudStorage/OneDrive-Personal/"
        ' Returned Paths:   https://d.docs.live.net  AND  OneDrive://
        ' Procedure:        Trim everything before the 4th '/'
    ' iCloud
        ' Local Paths:      "/Users/" & Environ("USER") & "/Library/Mobile Documents/com~apple~CloudDocs/"
        ' Returned Paths:   N/A
        ' Procedure:        No trim required. ThisWorkbook.Path returns full local path
    ' Google Drive
        ' Local Paths:      "/Users/" & Environ("USER") & "/Library/CloudStorage/GoogleDrive-[user]@gmail.com/"
        ' Returned Paths:   N/A
        ' Procedure:        No trim required. ThisWorkbook.Path returns full local path
    
    If Left(selectedPath, 23) = "https://d.docs.live.net" Or Left(selectedPath, 11) = "OneDrive://" Then
        For i = 1 To 4
            selectedPath = Mid(selectedPath, InStr(selectedPath, "/") + 1)
        Next
        
        #If Mac Then
            selectedPath = "/Users/" & Environ("USER") & "/Library/CloudStorage/OneDrive-Personal/" & selectedPath
        #Else
            selectedPath = Environ$("OneDrive") & "\" & Replace(selectedPath, "/", "\")
        #End If
    End If
End Sub

Private Function GetTempFilePath(ByVal fileName As String) As String
    #If Mac Then
        GetTempFilePath = Environ("TMPDIR") & fileName
    #Else
        GetTempFilePath = Environ("TEMP") & Application.PathSeparator & fileName
    #End If
End Function

Private Function MoveFile(ByVal initialPath As String, ByVal destinationPath As String) As Boolean
    Dim moveSuccessful As Boolean
    
    On Error Resume Next
    #If Mac Then
        If isAppleScriptInstalled Then
            moveSuccessful = AppleScriptTask(APPLE_SCRIPT_FILE, "CopyFile", initialPath & APPLE_SCRIPT_SPLIT_KEY & destinationPath)
        Else
            Name initialPath As destinationPath
            moveSuccessful = (Err.Number = 0)
        End If
    #Else
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFile initialPath, destinationPath
        moveSuccessful = (Err.Number = 0)
        Set fso = Nothing
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        If Not moveSuccessful Then
            Debug.Print "Failed to move template to " & destinationPath
            If Not isAppleScriptInstalled Then Debug.Print Err.Number & " - " & Err.Description
        End If
    #End If
    
    Err.Clear
    On Error GoTo 0
    MoveFile = moveSuccessful
End Function

Private Sub DeleteFile(ByVal filePath As String)
    #If Mac Then
        Dim appleScriptResult As Boolean
        
        If isAppleScriptInstalled Then
            appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", filePath)
            If appleScriptResult Then appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", filePath)
        Else
            Kill filePath
        End If
    #Else
        Dim fso As Object
        
        On Error Resume Next
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(filePath) Then fso.DeleteFile filePath, True
        
        filePath = Replace(filePath, " ", "%20")
        If fso.FileExists(filePath) Then fso.DeleteFile filePath, True
        
        Set fso = Nothing
        On Error GoTo 0
    #End If
End Sub

Private Sub DeleteExistingFolder(ByVal filePath As String)
    #If Mac Then
        Dim msgToDisplay As String, msgResult As Variant
        Dim scriptResult As Boolean

        If isAppleScriptInstalled Then
            scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "ClearFolder", filePath)
        Else
            msgToDisplay = "Because " & APPLE_SCRIPT_FILE & " is not installed, Excel is unable to delete any existing reports for this class. It is recommended to delete them before continuing." & _
                       vbNewLine & vbNewLine & "You can safely delete any files in '" & filePath & "' now and then click 'Okay' to continue."
            msgResult = DisplayMessage(msgToDisplay, vbOKOnly, "Notice")
        End If
    #Else
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")

        If Right(filePath, 1) = Application.PathSeparator Then
            filePath = Left(filePath, Len(filePath) - 1)
        End If

        fso.DeleteFolder filePath, True
        Set fso = Nothing
    #End If
End Sub

Private Function DisplayMessage(ByVal messageText As String, ByVal messageType As Integer, ByVal messageTitle As String, Optional ByVal dialogWidth As Integer = 250) As Variant
    On Error Resume Next
    #If Mac Then
        Dim dialogType As String, dialogResult As Variant, messageDisplayed As Boolean
        Dim lastError As String
        Dim i As Integer

        If isDialogToolkitInstalled Then
            Select Case messageType
                Case vbOKOnly, vbCritical, vbExclamation, vbInformation, vbQuestion
                    dialogType = "OkOnly"
                Case vbOKCancel
                    dialogType = "OkCancel"
                Case vbYesNo
                    dialogType = "YesNo"
                Case vbRetryCancel
                    dialogType = "RetryCancel"
                Case vbYesNoCancel
                    dialogType = "YesNoCancel"
                Case Else
                    dialogType = "OkOnly"
            End Select
            
            messageText = messageText & APPLE_SCRIPT_SPLIT_KEY & dialogType & APPLE_SCRIPT_SPLIT_KEY & messageTitle & APPLE_SCRIPT_SPLIT_KEY & dialogWidth
            dialogResult = ""
            
            On Error Resume Next
            Do While Not messageDisplayed
                dialogResult = AppleScriptTask("DialogDisplay.scpt", "DisplayDialog", messageText)
                messageDisplayed = (dialogResult <> "")
                i = i + 1
                
                #If PRINT_DEBUG_MESSAGES Then
                    If Err.Number <> 0 Then lastError = Err.Number & " - " & Err.Description
                #End If
                
                If i >= 30 Then
                    DisplayMessage = MsgBox(messageText, messageType, messageTitle)
                    Exit Do
                End If
            Loop
            On Error GoTo 0
            #If PRINT_DEBUG_MESSAGES Then
                If lastError = "" Then lastError = "N/A"
                Debug.Print "Number of attempts: " & i & vbNewLine & "Final error: " & lastError
            #End If
            Exit Function
        End If
    #End If
    On Error GoTo 0
    
    DisplayMessage = MsgBox(messageText, messageType, messageTitle)
End Function
