Option Explicit

#Const Windows = (Mac = 0)

Public Type StudentRecords
    englishName         As String
    koreanName          As String
    GrammarScore        As String
    PronunciationScore  As String
    FluencyScore        As String
    MannerScore         As String
    ContentScore        As String
    EffortScore         As String
    comment             As String
    OverallGrade        As String
End Type

Public Type ClassRecords
    englishTeacher      As String
    KoreanTeacher       As String
    classLevel          As String
    classDays           As String
    classTime           As String
    EvalulationDate     As String
End Type

Public Type CertificateDesign
    Type                As String
    Layout              As String
    Design              As String
    borderType          As String
    BorderColor         As String
    borderColorCode     As String
End Type

Public Type WinningStudents
    englishName         As String
    koreanName          As String
End Type

Public Sub CreateReportsAndCertificates(ByRef ws As Worksheet, ByVal clickedButton As String)
    Const ERR_RESOURCES_FOLDER          As String = "Display.ErrorMessages.UnableToCreateResourcesFolder"
    Const ERR_FONT_INSTALLATION         As String = "Display.ErrorMessages.FontInstallationFailed"
    Const ERR_INCOMPLETE_RECORDS        As String = "Display.ErrorMessages.IncompleteRecords"
    Const ERR_INCOMPLETE_WINNERS_LIST   As String = "Display.ErrorMessages.IncompleteWinnersList"
    Const ERR_LOADING_POWERPOINT        As String = "Display.ErrorMessages.ErrorOpeningPowerPoint"
    Const ERR_LOADING_TEMPLATE          As String = "Display.ErrorMessages.ErrorOpeningTemplate"
    Const ERR_MISSING_SHAPES            As String = "Display.ErrorMessages.MissingShapes"
    Const MSG_SAVE_FAILED               As String = "Display.ErrorMessages.ErrorSavingFile"
    Const MSG_ZIP_FAILED                As String = "Display.ErrorMessages.ErrorZippingFiles"
    Const MSG_SUCCESS                   As String = "Display.Reports.Success"
    Const ERR_INVALID_BUTTON            As String = "Display.Reports.InvalidReportButton"
    
    Dim pptApp           As Object
    Dim pptDoc           As Object
    Dim classInformation As ClassRecords
    Dim studentData()    As StudentRecords
    Dim winnerValues()   As WinningStudents
    Dim generateProcess  As String
    Dim pathToOpen       As String
    Dim resourcesFolder  As String
    Dim resultMsg        As String
    Dim saveResult       As Boolean

    RepairLayouts ws

    generateProcess = DetermineGenerationProcess(clickedButton)
    classInformation = LoadClassInformation(ws.Range(g_CLASS_INFO).Value)
    
    Select Case generateProcess
        Case "Reports", "Proofs"
            If Not GenerateAndVerifyRecordsAreComplete(ws, classInformation, studentData()) Then
                resultMsg = ERR_INCOMPLETE_RECORDS
                GoTo Cleanup
            End If
            
        Case "Certificates"
            If WinnersListsIncomplete(ws) Then
                resultMsg = ERR_INCOMPLETE_WINNERS_LIST
                GoTo Cleanup
            End If
            winnerValues() = LoadWinnerValues(ws)
        Case Else
            resultMsg = ERR_INVALID_BUTTON
            GoTo Cleanup
    End Select

    ' Possibly move and/or add this to WorkbookOpen
    ' This will then lead to a rewrite of filepath code
    Application.DefaultFilePath = GetDefaultFolderPaths("Base")

    resourcesFolder = GetDefaultFolderPaths("Resources")
    If Not CheckForAndAttemptToCreateFolder(resourcesFolder) Then
        resultMsg = ERR_RESOURCES_FOLDER
        GoTo Cleanup
    End If

    Select Case generateProcess
        Case "Reports", "Proofs"
            saveResult = CreateReportFiles(ws, classInformation, studentData(), generateProcess)
        Case "Certificates"
            saveResult = CreateCertificateFiles(ws, winnerValues(), classInformation)
    End Select

    If Not saveResult Then
        resultMsg = MSG_SAVE_FAILED
        GoTo Cleanup
    End If
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.FileManagement.FileCreationSaveComplete")
    End If
    
    resultMsg = MSG_SUCCESS
    
    If generateProcess = "Reports" Then
        Select Case False
            Case g_UserOptions.ZipSupportEnabled, Verify7ZipIsPresent
                g_UserOptions.ZipSupportEnabled = CheckFor7Zip(True)
                ToggleSheetProtection Options, False
                WriteNewRangeValue Options.Range(g_7ZIP_SUPPORT_STATUS), IIf(g_UserOptions.ZipSupportEnabled, "Yes", "No")
                ToggleSheetProtection Options, True
        End Select
        
        If g_UserOptions.ZipSupportEnabled Then
            saveResult = ZipReports(ws, resourcesFolder, classInformation)
        End If
        
        If Not saveResult Then
            resultMsg = MSG_ZIP_FAILED
            GoTo Cleanup
        End If
    End If
    
    If g_UserOptions.OpenSavePathWhenDone And saveResult Then
        pathToOpen = SetSavePath(ws, generateProcess)
        OpenDestinationFolder pathToOpen
    End If
    
Cleanup:
    If resultMsg <> vbNullString Then DisplayMessage resultMsg
    
    ' Final call to ensure the static variable gets reset
    SetSavePath ws, generateProcess, True

    If Not pptDoc Is Nothing Or Not pptApp Is Nothing Then
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.CodeExecution.BeginningCleanup")
        End If
        
        If Not pptDoc Is Nothing Then
            CloseTemplate pptDoc
        End If
        
        If Not pptApp Is Nothing Then
            ClosePowerPoint pptApp
        End If
    End If
End Sub

Private Function DetermineGenerationProcess(ByVal clickedButton As String) As String
    ' Button names are prefixed with "Button_Generate", so works fine
    DetermineGenerationProcess = Mid$(clickedButton, 16, Len(clickedButton) - 15)
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.FileManagement.FileCreationStart", DetermineGenerationProcess)
    End If
End Function

Public Function GenerateAndVerifyRecordsAreComplete(ByRef ws As Worksheet, ByRef classInformation As ClassRecords, ByRef studentData() As StudentRecords) As Boolean
    Dim invalidCategory As String
    Dim invalidRecordIndex As Long
    Dim finalRow As Long

    If Not ValidateClassData(classInformation, invalidCategory) Then
        ' Display message to inform user where the problem is
        GenerateAndVerifyRecordsAreComplete = False
        Exit Function
    End If

    finalRow = DetermineFinalRowForClass(ws)

    If Not SheetContainsStudentRecords(finalRow) Then
        DisplayMessage "Display.StudentRecords.NoStudentsFound"
        
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.StudentRecords.NoStudentsFound")
        End If

        GenerateAndVerifyRecordsAreComplete = False
        Exit Function
    End If

    If Not VerifyRecordsAreComplete(ws, finalRow) Then
        DisplayMessage "Display.StudentRecords.ErrorReadingData"

        GenerateAndVerifyRecordsAreComplete = False
        Exit Function
    End If

    studentData() = LoadStudentRecords(ws, finalRow)

    If Not ValidateStudentData(studentData(), invalidRecordIndex, invalidCategory) Then
        'Display error to inform where problem exists

        GenerateAndVerifyRecordsAreComplete = False
        Exit Function
    End If

    GenerateAndVerifyRecordsAreComplete = True
End Function

Private Function LoadClassInformation(ByVal classInfoValue As Variant) As ClassRecords
    Dim classInformation As ClassRecords
    
    With classInformation
        .englishTeacher = Trim$(classInfoValue(1, 1))
        .KoreanTeacher = Trim$(classInfoValue(2, 1))
        .classLevel = Trim$(classInfoValue(3, 1))
        .classDays = Trim$(classInfoValue(4, 1))
        .classTime = Trim$(classInfoValue(5, 1))
        .EvalulationDate = Trim$(CStr(classInfoValue(6, 1)))
    End With
    
    LoadClassInformation = classInformation
End Function

Public Function DetermineFinalRowForClass(ByRef ws As Worksheet) As Long
    Const ENGLISH_NAME_COLUMN As Long = 2 ' Column B
    DetermineFinalRowForClass = ws.Cells.Item(ws.Rows.Count, ENGLISH_NAME_COLUMN).End(xlUp).Row
End Function

Private Function WinnersListsIncomplete(ByRef ws As Worksheet) As Boolean
    Dim winnersList As Variant
    Dim incompleteWinnersList As Boolean
    
    winnersList = ws.Range(g_WINNER_NAMES).Value
    
    incompleteWinnersList = (winnersList(1, 1) = vbNullString) Or _
                            (winnersList(2, 1) = vbNullString And winnersList(3, 1) <> vbNullString)
    
    If incompleteWinnersList Then
        If DisplayMessage("Display.StudentRecords.IncompleteWinnersList") = vbYes Then
            UpdateWinnersLists ws, True
            incompleteWinnersList = False
        End If
    End If
    
    WinnersListsIncomplete = incompleteWinnersList
End Function

Public Function LoadStudentRecords(ByRef ws As Worksheet, Optional ByVal finalRow As Long = 0) As StudentRecords()
    Dim studentData() As StudentRecords
    Dim currentStudentData As Variant
    Dim totalNumberOfRecords As Long
    Dim i As Long
    
    If finalRow = 0 Then
        finalRow = DetermineFinalRowForClass(ws)
    End If
    
    totalNumberOfRecords = finalRow - 7

    ReDim studentData(1 To totalNumberOfRecords)

    For i = g_FIRST_STUDENT_ROW To finalRow
        currentStudentData = ws.Range("B" & CStr(i) & ":J" & CStr(i)).Value
        LoadStudentData currentStudentData, studentData(i - g_STUDENT_INDEX_OFFSET)
    Next i
    
    LoadStudentRecords = studentData()
End Function

Private Function LoadStudentData(ByVal studentData As Variant, ByRef studentScores As StudentRecords) As Boolean
    With studentScores
        .englishName = Trim$(studentData(1, 1))
        .koreanName = Trim$(studentData(1, 2))
        .GrammarScore = FormatStudentScoreWhenLoaded(studentData(1, 3))
        .PronunciationScore = FormatStudentScoreWhenLoaded(studentData(1, 4))
        .FluencyScore = FormatStudentScoreWhenLoaded(studentData(1, 5))
        .MannerScore = FormatStudentScoreWhenLoaded(studentData(1, 6))
        .ContentScore = FormatStudentScoreWhenLoaded(studentData(1, 7))
        .EffortScore = FormatStudentScoreWhenLoaded(studentData(1, 8))
        .comment = Trim$(studentData(1, 9))
    End With
End Function

Private Function FormatStudentScoreWhenLoaded(ByVal dataToFormat As String) As String
    FormatStudentScoreWhenLoaded = Trim$(UCase$(CStr(dataToFormat)))
End Function

Private Function LoadWinnerValues(ByRef ws As Worksheet) As WinningStudents()
    Dim topThreeStudents As Variant
    Dim winningNameHalves() As String
    Dim tmpList() As WinningStudents
    Dim numberOfWinners As Long
    Dim i As Long

    topThreeStudents = ws.Range(g_WINNER_NAMES).Value

    For i = 1 To 3
        If topThreeStudents(i, 1) = vbNullString Then
            Exit For
        End If

        numberOfWinners = i
    Next i

    If numberOfWinners = 0 Then
        ' Figure out some kind of check value to return
    End If

    ReDim tmpList(1 To numberOfWinners)
    For i = 1 To numberOfWinners
        ' SplitWinnerName topThreeStudents(i, 1) tmpList(i).englishName, tmpList(i).koreanName
        winningNameHalves() = Split(topThreeStudents(i, 1), " (")

        tmpList(i).englishName = Trim$(winningNameHalves(0))
        tmpList(i).koreanName = Trim$(Left$(winningNameHalves(1), Len(winningNameHalves(1)) - 1))
    Next i

    LoadWinnerValues = tmpList()
End Function

Private Function CloseTemplateIfOpen(ByRef pptApp As Object, ByVal templatePath As String, ByVal templateToUse As String) As Boolean
    Dim pptOpenDoc As Object
    Dim pptToClose As Object
    Dim openDocName As String
    Dim tempateAlreadyOpen As Boolean
    Dim otherPptDocIsOpen As Boolean

    On Error Resume Next
    For Each pptOpenDoc In pptApp.Presentations
        openDocName = ConvertToLocalPath(pptOpenDoc.fullName)

        If StrComp(templatePath, openDocName, vbTextCompare) = 0 Or StrComp(templateToUse, openDocName, vbTextCompare) = 0 Then
            tempateAlreadyOpen = True
            Set pptToClose = pptOpenDoc
        Else
            otherPptDocIsOpen = True
        End If

        If tempateAlreadyOpen And otherPptDocIsOpen Then Exit For
    Next pptOpenDoc
    On Error GoTo 0

    If otherPptDocIsOpen Then
        ' Warn the user and cancel or wait
    End If

    If tempateAlreadyOpen Then
        ' Display a warning?
        ' This isn't working
        On Error Resume Next
        pptToClose.Close
        If Err.Number = 0 Then Set pptToClose = Nothing
        On Error GoTo 0
    End If

    CloseTemplateIfOpen = (pptToClose Is Nothing)
End Function

Private Function VerifyPptDocIsClosed(ByRef pptApp As Object, ByVal pathToClosedDoc As String) As Boolean
    Dim openPptDoc As Object
    Dim fileIsClosed As Boolean
    
    fileIsClosed = True
    
    On Error Resume Next
    For Each openPptDoc In pptApp.Presentations
        If StrComp(ConvertToLocalPath(openPptDoc.fullName), pathToClosedDoc, vbTextCompare) = 0 Then
            fileIsClosed = False
            Exit For
        End If
    Next openPptDoc
    On Error GoTo 0
    
    VerifyPptDocIsClosed = fileIsClosed
End Function

Private Function LoadPowerPoint(ByRef pptApp As Object) As Boolean
    Dim isOpen As Boolean
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.PowerPoint.Opening")
    End If
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.PowerPoint.CheckingForOpenInstances")
    End If

    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
    
    isOpen = Not pptApp Is Nothing

    If Not isOpen Then
        isOpen = OpenPowerPoint(pptApp)
    End If

    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.PowerPoint.LoadedStatus", isOpen)
    End If
    
    LoadPowerPoint = isOpen
End Function

Private Function OpenPowerPoint(ByRef pptApp As Object) As Boolean
    On Error Resume Next
#If Mac Then
    Dim appleScriptResult As String

    appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "LoadApplication", "Microsoft PowerPoint")

    If g_UserOptions.EnableLogging Then
        If appleScriptResult <> vbNullString Then
            DebugAndLogging appleScriptResult
        End If
    End If

    appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "IsAppLoaded", "Microsoft PowerPoint")
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging INDENT_LEVEL_1 & appleScriptResult
    End If
    
    Set pptApp = GetObject(, "PowerPoint.Application")

    If Err.Number <> 0 Then
        DisplayMessage "Display.PowerPoint.ErrorLoadingViaAppleScript", Err.Number, Err.Description, appleScriptResult
    End If
#Else
    Set pptApp = CreateObject("PowerPoint.Application")

    ' If err.Number <> 0 Then
    '     DisplayMessage "Display.PowerPoint.ErrorLoadingViaAppleScript", err.Number, err.Description, appleScriptResult
    ' End If
#End If
    On Error GoTo 0

    OpenPowerPoint = Not pptApp Is Nothing
End Function

Private Function LoadTemplate(ByRef pptApp As Object, ByRef pptDoc As Object, ByVal generateProcess As String) As Boolean
    Const CERTIFICATE_TEMPLATE_FULL As String = "CertificateTemplate-Full"
    Const CERTIFICATE_TEMPLATE_LITE As String = "CertificateTemplate-Lite"
    Const REPORT_TEMPLATE_FULL      As String = "SpeakingEvaluationTemplate-Full"
    Const REPORT_TEMPLATE_LITE      As String = "SpeakingEvaluationTemplate-Lite"
    
    Dim subfolderPath As String
    Dim templatePath As String
    Dim templateType As String
    Dim tempateToOpen As String
    Dim templateName As String
    Dim isLoaded As Boolean
    Dim autoSaveDisabled As Boolean

    Select Case generateProcess
        Case "Reports", "Proofs"
            templateType = IIf(g_UserOptions.AllFontsAreInstalled, REPORT_TEMPLATE_LITE, REPORT_TEMPLATE_FULL)
        Case "Certificates"
            templateType = IIf(g_UserOptions.AllFontsAreInstalled, CERTIFICATE_TEMPLATE_LITE, CERTIFICATE_TEMPLATE_FULL)
    End Select
    
    templateName = ReadValueFromDictionary(g_dictFileData, templateType, "filename")
    templatePath = GetDefaultFolderPaths("Resources") & templateName
    tempateToOpen = PrepareRequiredFile(templateName, templatePath, templateType)

    If tempateToOpen = vbNullString Then
        LoadTemplate = False
        Exit Function
    End If

    If Not CloseTemplateIfOpen(pptApp, templatePath, tempateToOpen) Then
        ' Possibly display an message asking if they wish to continue?
        LoadTemplate = False
        Exit Function
    End If

    On Error Resume Next
    Set pptDoc = pptApp.Presentations.Open(tempateToOpen)
    On Error GoTo 0

    isLoaded = Not pptDoc Is Nothing

    If isLoaded Then
        DisableAutoSave pptApp, pptDoc
    End If

    LoadTemplate = isLoaded
End Function

Private Function CreateReportFiles(ByRef ws As Worksheet, ByRef classInformation As ClassRecords, ByRef studentData() As StudentRecords, ByVal generateProcess As String) As Boolean
    Const SIG_SHAPE_NAME      As String = "mySignature-Enabled"
    Const SIG_EMBEDDED_TOGGLE As String = "Button_SignatureEmbedded"
    
    Dim pptApp   As Object
    Dim pptDoc   As Object
    Dim fileName As String
    Dim i        As Long
    
    If Not LoadPowerPoint(pptApp) Then
        CreateReportFiles = False
        Exit Function
    End If

    If Not LoadTemplate(pptApp, pptDoc, generateProcess) Then
        CreateReportFiles = False
        Exit Function
    End If

    SetPowerPointViewSettings pptApp

    If Not GenerateReportSharedHeader(pptDoc, classInformation) Then
        CreateReportFiles = False
        Exit Function
    End If

    On Error Resume Next
    If Options.Shapes(SIG_EMBEDDED_TOGGLE).Visible = msoTrue Then
        If pptDoc.Slides(1).Shapes(SIG_SHAPE_NAME) Is Nothing Then
            InsertSignature pptDoc, SIG_SHAPE_NAME
        End If
    End If
    On Error GoTo 0

    For i = LBound(studentData) To UBound(studentData)
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.Reports.GeneratingFile", i, UBound(studentData))
        End If
        
        If Not GenerateReportBody(pptDoc, studentData(i)) Then
            CreateReportFiles = False
            Exit Function
        End If
        
        fileName = studentData(i).koreanName & "(" & SanitizeFileName(studentData(i).englishName) & ")" & " - " & classInformation.classDays

        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.Reports.Filename", fileName)
        End If

        If Not WriteFileToDisk(ws, pptApp, pptDoc, generateProcess, fileName) Then
            CreateReportFiles = False
            Exit Function
        End If
    Next i
    
    CloseTemplate pptDoc
    ClosePowerPoint pptApp

    CreateReportFiles = True
End Function

Private Function GenerateReportSharedHeader(ByRef pptDoc As Object, ByRef classInformation As ClassRecords) As Boolean
    On Error GoTo ErrorHandler
    With pptDoc.Slides(1).Shapes("Report_Header").GroupItems
        WriteValuesToReport .Item("Native_Teacher"), classInformation.englishTeacher, "Just Another Hand"
        WriteValuesToReport .Item("Korean_Teacher"), classInformation.KoreanTeacher, "Kakao Big Sans"
        WriteValuesToReport .Item("Grade_Level"), classInformation.classLevel, "Just Another Hand"
        WriteValuesToReport .Item("Eval_Date"), Format$(CDate(classInformation.EvalulationDate), "DD MMM. YYYY"), "Just Another Hand"
    End With
    On Error GoTo 0

    GenerateReportSharedHeader = True
    Exit Function
ErrorHandler:
    ' Display an error

    If g_UserOptions.EnableLogging Then
        ' Output an error to the log
    End If

    GenerateReportSharedHeader = False
End Function

Private Function GenerateReportBody(ByRef pptDoc As Object, ByRef studentData As StudentRecords) As Boolean
    Dim englishNameTextboxHeight    As Long
    Dim englishNameTextboxTop       As Long
    Dim englishNameFontSize         As Long
    
    Dim i As Long
    
    ReformatEnglishName studentData.englishName, englishNameTextboxHeight, englishNameTextboxTop, englishNameFontSize

    With pptDoc.Slides(1).Shapes
        With .Item("Report_Header").GroupItems
            With .Item("English_Name")
                .Height = englishNameTextboxHeight
                .Top = englishNameTextboxTop
            End With

            WriteValuesToReport .Item("English_Name"), studentData.englishName, "Just Another Hand", englishNameFontSize
            WriteValuesToReport .Item("Korean_Name"), studentData.koreanName, "Kakao Big Sans"
        End With

        WriteValuesToReport .Item("Comments"), studentData.comment, "Just Another Hand", 24
        .Item("Comments").TextFrame2.AutoSize = msoAutoSizeTextToFitShape
    End With

    ToggleScoreVisibility pptDoc, studentData
    
    ' Add error handling
    GenerateReportBody = True
End Function

Private Function CreateCertificateFiles(ByRef ws As Worksheet, ByRef winnerValues() As WinningStudents, ByRef classInformation As ClassRecords) As Boolean
    Dim pptApp As Object
    Dim pptDoc As Object
    
    Dim certificateSettings As CertificateDesign
    Dim elementPositioning As String
    Dim fileName As String
    
    Dim studentRanking As Long
    Dim i As Long
    Dim j As Long
    
    If Not LoadPowerPoint(pptApp) Then
        CreateCertificateFiles = False
        Exit Function
    End If

    If Not LoadTemplate(pptApp, pptDoc, "Certificates") Then
        CreateCertificateFiles = False
        Exit Function
    End If
    
    certificateSettings = LoadCertificateDesign()
    elementPositioning = GetElementPositioning(certificateSettings.Design)
    
    ToggleCertificateStyleShapes pptDoc, certificateSettings.borderType, certificateSettings.borderColorCode, elementPositioning
    ImportClassDataToCertificates pptDoc, elementPositioning, classInformation.englishTeacher, classInformation.classLevel, classInformation.EvalulationDate
    ImportHeaderToCertificates pptDoc, elementPositioning, certificateSettings.Type

    For i = LBound(winnerValues) To UBound(winnerValues)
        studentRanking = i + 1

        ToggleStudentPlacementShapes pptDoc, elementPositioning, certificateSettings.Design, studentRanking
        ImportStudentSpecificDataToCertificates pptDoc, elementPositioning, winnerValues(i).englishName, winnerValues(i).koreanName
        

        fileName = winnerValues(i).koreanName & "(" & winnerValues(i).englishName & ") - " & classInformation.classDays

        If Not WriteFileToDisk(ws, pptApp, pptDoc, "Certificates", fileName) Then
            CreateCertificateFiles = False
            Exit Function
        End If
    Next i

    CloseTemplate pptDoc
    ClosePowerPoint pptApp

    CreateCertificateFiles = True
End Function

Public Function LoadCertificateDesign() As CertificateDesign
    Dim certSettings() As Variant
    Dim tmpData As CertificateDesign
    
    certSettings = Options.Range(g_CERTIFICATE_OPTIONS).Value
    
    tmpData.Type = certSettings(1, 1)
    tmpData.Layout = certSettings(2, 1)
    tmpData.Design = certSettings(3, 1)
    tmpData.borderType = certSettings(4, 1)
    tmpData.BorderColor = certSettings(5, 1)
    tmpData.borderColorCode = certSettings(6, 1)

    LoadCertificateDesign = tmpData
End Function

Private Sub ImportHeaderToCertificates(ByRef pptDoc As Object, ByVal elementPositioning As String, ByVal headerType As String)
    Dim headerFirstLine As String
    Dim headerSecondLine As String
    Dim headerFirstLineFontSize As Long
    Dim headerSecondLineFontSize As Long
    
    With pptDoc.Slides(1).Shapes("Title_" & elementPositioning).TextFrame.TextRange
        Select Case headerType
            Case "Speech Contest"
                headerFirstLine = "SPEECH CONTEST"
                headerSecondLine = "WINNER"
                headerFirstLineFontSize = 38
                headerSecondLineFontSize = 84
            Case "Winter Speeches", "Spring Speeches", "Summer Speeches", "Fall Speeches", "Autumn Speeches"
                headerFirstLine = UCase$(Split(headerType)(0) & " SPEAKING")
                headerSecondLine = "EVALUATIONS"
                headerFirstLineFontSize = 38
                headerSecondLineFontSize = 60
        End Select
        
        .text = headerFirstLine & vbNewLine & headerSecondLine
        .Characters.Font.Name = "Constantia"
        .Characters(1, Len(headerFirstLine)).Font.Size = headerFirstLineFontSize
        .Characters(Len(headerFirstLine) + 2, Len(headerSecondLine)).Font.Size = headerSecondLineFontSize
    End With
End Sub

Private Sub ImportClassDataToCertificates(ByRef pptDoc As Object, ByVal elementPositioning As String, ByVal englishTeacher As String, ByVal classLevel As String, ByVal evalDate As String)
    With pptDoc.Slides(1).Shapes
        With .Item("Modified_Elements").GroupItems
            .Item("Teacher_" & elementPositioning).TextFrame.TextRange.text = "in " & englishTeacher & " Teacher's"
            .Item("Date_Text_" & elementPositioning).TextFrame.TextRange.text = evalDate
        End With
        .Item("Levels").GroupItems(classLevel & "_" & elementPositioning).Visible = msoTrue
    End With
End Sub

Private Sub ImportStudentSpecificDataToCertificates(ByRef pptDoc As Object, ByVal elementPositioning As String, ByVal englishName As String, ByVal koreanName As String)
    Dim fullName As String
    Dim koreanTextLength As Long
    
    fullName = koreanName
    If Len(englishName) < 10 Then
        fullName = fullName & " (" & englishName & ")"
    End If

    koreanTextLength = Len(koreanName)
    With pptDoc.Slides(1).Shapes("Modified_Elements").GroupItems("Student_" & elementPositioning).TextFrame.TextRange
        .text = fullName
        .Characters(1, koreanTextLength).Font.Name = "Kakao Big Sans"
        .Characters(koreanTextLength + 1, Len(fullName) - koreanTextLength).Font.Name = "Constantia"
    End With
End Sub

Private Sub CloseTemplate(ByRef pptDoc As Object)
    pptDoc.Close
    Set pptDoc = Nothing
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.PowerPoint.ClosingOpenTemplate", (pptDoc Is Nothing))
    End If
End Sub

Private Sub ClosePowerPoint(ByRef pptApp As Object)
    If Not pptApp Is Nothing Then
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.PowerPoint.AttemptingToClose")
        End If

        pptApp.Quit
        Set pptApp = Nothing

        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.CodeExecution.Result", INDENT_LEVEL_1, (pptApp Is Nothing))
        End If
    End If

#If Mac Then
    Dim closeResult As String
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.PowerPoint.AttemptingToCloseExtraMacStep")
    End If

    closeResult = AppleScriptTask(APPLE_SCRIPT_FILE, "ClosePowerPoint", vbNullString)

    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.CodeExecution.Result", INDENT_LEVEL_1, closeResult)
    End If
#End If
End Sub

Private Function ZipReports(ByRef ws As Worksheet, ByVal resourcesFolder As String, ByRef classInformation As ClassRecords) As Boolean
    Dim savePath    As String
    Dim zipCommand  As String
    Dim zipName     As String
    Dim zipPath     As String
    Dim saveResult  As Boolean

    savePath = SetSavePath(ws)
    EnsureTrailingPathSeparator savePath

    With classInformation
        zipName = .classLevel & " (" & .KoreanTeacher & " - " & _
                  .classDays & " " & .classTime & ").zip"
    End With

    zipPath = savePath & zipName

    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.Zip.ZippingReport", zipName, savePath)
    End If
    
    zipCommand = CreateZipCommand(resourcesFolder, zipPath, savePath, "*.pdf")

    saveResult = CreateZipFile(zipCommand, savePath, zipName)

    If saveResult Then
        RemovePdfReports savePath

        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.Zip.Successful")
        End If
    Else
        If g_UserOptions.EnableLogging Then
            ' This will need to be updated
            DebugAndLogging GetMsg("Debug.Zip.Failed", Err.Number, Err.Description)
        End If
    End If

    ZipReports = saveResult
End Function

Private Function CreateZipFileName(ByVal zipNameElementsRng As Range) As String
    Dim zipFileNameElements As Variant
    
    zipFileNameElements = zipNameElementsRng.Value
    CreateZipFileName = zipFileNameElements(2, 1) & " (" & zipFileNameElements(1, 1) & " - " & _
                        zipFileNameElements(3, 1) & " " & zipFileNameElements(4, 1) & ").zip"
End Function

Private Function CreateZipFilePath(ByVal savePath As Variant, ByVal zipFileName As String) As String
#If Mac Then
    CreateZipFilePath = savePath & zipFileName
#Else
    CreateZipFilePath = GetDefaultFolderPaths("Temp") & zipFileName
#End If
End Function

Private Function CreateZipCommand(ByVal resourcesFolder As String, ByVal zipPath As String, ByVal savePath As String, ByVal fileExtension As String) As String
#If Mac Then
    Const ZIP_TOOL_BINARY As String = "7zz"
#Else
    Const ZIP_TOOL_BINARY As String = "7za.exe"
#End If
    Const ZIP_TOOL_FLAGS  As String = " a " ' Leading and trailing spaces added here for simplicity
    Const ZIP_PATH_SPACER As String = " "
    
    Dim zipToolPath As String

    zipToolPath = resourcesFolder & Application.PathSeparator & ZIP_TOOL_BINARY
    
    CreateZipCommand = Chr(34) & _
                       zipToolPath & Chr(34) & _
                       ZIP_TOOL_FLAGS & Chr(34) & _
                       zipPath & Chr(34) & _
                       ZIP_PATH_SPACER & Chr(34) & _
                       savePath & _
                       fileExtension & _
                       Chr(34)
End Function

Private Function GetArchiverPath(ByVal resourcesFolder As String) As String
#If Mac Then
    GetArchiverPath = resourcesFolder & Application.PathSeparator & "7zz"
#Else
    GetArchiverPath = resourcesFolder & Application.PathSeparator & "7za.exe"
#End If
End Function

Private Function CreateZipFile(ByVal zipCommand As String, ByVal savePath As String, ByVal zipName As String) As Boolean
#If Mac Then
    Dim scriptResultString As String
    
    On Error Resume Next
    scriptResultString = AppleScriptTask(APPLE_SCRIPT_FILE, "CreateZipWithLocal7Zip", zipCommand)
    On Error GoTo 0
    
    CreateZipFile = (scriptResultString = "Success")
#Else
    DeleteFile savePath & zipName
    
    Shell zipCommand, vbHide
    
    Application.Wait Now + TimeValue("0:00:02")
    
    CreateZipFile = DoesFileExist(savePath & zipName)
#End If
End Function

Private Sub RemovePdfReports(ByVal savePath As String)
#If Mac Then
    Dim scriptResult As Boolean
    
    On Error Resume Next
    scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "ClearPDFsAfterZipping", savePath)
    On Error GoTo 0
#Else
    Dim fso As Object
    Dim objFile As Object
    Dim objFolder As Object

    EnsureTrailingPathSeparator savePath
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objFolder = fso.GetFolder(savePath)
    
    On Error Resume Next
    For Each objFile In objFolder.Files
        If LCase$(fso.GetExtensionName(objFile.Name)) = "pdf" Then
            objFile.Delete True
            If g_UserOptions.EnableLogging Then
                If Err.Number <> 0 Then
                    DebugAndLogging GetMsg("Debug.FileManagement.DeletingFileFailed", objFile.Name, Err.Number, Err.Description)
                    Err.Clear
                End If
            End If
        End If
    Next objFile
    On Error GoTo 0
#End If
End Sub

Private Sub WriteValuesToReport(ByRef targetBox As Object, ByVal textboxValue As String, ByVal desiredFont As String, Optional ByVal fontSize As Long = 0, Optional ByVal boldFont As Boolean = False)
    If fontSize = 0 Then
        fontSize = IIf(desiredFont = "Just Another Hand", 20, 16)
    End If
    
    With targetBox.TextFrame.TextRange
        If .text <> textboxValue Then .text = textboxValue
        With .Font
            If .Name <> desiredFont Then .Name = desiredFont
            If .Size <> fontSize Then .Size = fontSize
            If .Bold <> boldFont Then .Bold = boldFont
        End With
    End With
End Sub

Private Sub ReformatEnglishName(ByRef englishName As String, ByRef englishNameTextboxHeight As Long, ByRef englishNameTextboxTop As Long, ByRef englishNameFontSize As Long)
    Dim textLength As Long
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim posToReplace As Long
    Dim i As Long
    
    textLength = Len(englishName)
    upperBound = CInt(textLength * 0.8)
    
    Select Case textLength
        Case Is > 35
            lowerBound = 21
        Case Is > 30
            lowerBound = 19
        Case Is > 25
            lowerBound = 17
        Case Else
            lowerBound = 15
    End Select
    
    ' Step 1: Set textbox customizations
    If textLength > 20 Then
        englishNameTextboxHeight = 40
        englishNameTextboxTop = 64
        englishNameFontSize = 14
    Else
        englishNameTextboxHeight = 28
        englishNameTextboxTop = 78
        englishNameFontSize = 20
        Exit Sub
    End If
    
    ' Step 2: Locate spaces with englishName
    posToReplace = -1
    For i = lowerBound To upperBound
        If Mid$(englishName, i, 1) = " " Then
            posToReplace = i
            Exit For
        End If
    Next i
    
    ' Step 3: Update englishName
    If posToReplace <> -1 Then
        ' Step 3a: Replace a space with vbNewLine to split the string into two lines
        englishName = Left$(englishName, posToReplace - 1) & vbNewLine & Mid$(englishName, posToReplace + 1)
    Else
        ' Step 3b: Hyphenate and split into two lines
        englishName = Left$(englishName, lowerBound) & "-" & vbNewLine & Mid$(englishName, lowerBound + 1)
    End If
End Sub

Public Function CalculateOverallGrade(ByRef studentScores As StudentRecords) As String
    Dim gradeValues() As Variant
    Dim calculatedScore As Double
    
    gradeValues() = Array("C", "B", "B+", "A", "A+")
    
    With studentScores
        calculatedScore = ConvertGradeToPoints(.GrammarScore, gradeValues())
        calculatedScore = calculatedScore + ConvertGradeToPoints(.PronunciationScore, gradeValues())
        calculatedScore = calculatedScore + ConvertGradeToPoints(.FluencyScore, gradeValues())
        calculatedScore = calculatedScore + ConvertGradeToPoints(.MannerScore, gradeValues())
        calculatedScore = calculatedScore + ConvertGradeToPoints(.ContentScore, gradeValues())
        calculatedScore = calculatedScore + ConvertGradeToPoints(.EffortScore, gradeValues())
    End With

    calculatedScore = calculatedScore / 6

    If calculatedScore - Int(calculatedScore) >= 0.4 Then
        calculatedScore = Int(calculatedScore)
    Else
        calculatedScore = Int(calculatedScore) - 1
    End If

    CalculateOverallGrade = gradeValues(calculatedScore)
End Function

Private Function ConvertGradeToPoints(ByVal score As String, ByRef gradeValues() As Variant) As Long
    Dim i As Long
    
    For i = LBound(gradeValues) To UBound(gradeValues)
        If gradeValues(i) = score Then
            ConvertGradeToPoints = i + 1
            Exit Function
        End If
    Next i
End Function

Private Sub ToggleScoreVisibility(ByRef pptDoc As Object, ByRef studentData As StudentRecords)
    Dim scoreValues     As Variant
    Dim scoreCategories As Variant
    Dim currentCategory As String
    Dim currentScore    As String
    Dim i               As Long

    scoreCategories = Array("Grammar_", "Pronunciation_", "Fluency_", "Manner_", "Content_", "Effort_", "Result_")
    scoreValues = Array(studentData.GrammarScore, studentData.PronunciationScore, studentData.FluencyScore, _
                        studentData.MannerScore, studentData.ContentScore, studentData.EffortScore, studentData.OverallGrade)
    
    For i = LBound(scoreCategories) To UBound(scoreCategories)
        currentCategory = scoreCategories(i)
        currentScore = scoreValues(i)
        
        With pptDoc.Slides(1).Shapes(currentCategory & "Scores").GroupItems
            .Item(currentCategory & "A+").Visible = (currentScore = "A+")
            .Item(currentCategory & "A").Visible = (currentScore = "A")
            .Item(currentCategory & "B+").Visible = (currentScore = "B+")
            .Item(currentCategory & "B").Visible = (currentScore = "B")
            .Item(currentCategory & "C").Visible = (currentScore = "C")
        End With
    Next i
End Sub

Private Sub InsertSignature(ByRef pptDoc As Object, ByVal sigShapeName As String)
    Dim sigShape As Object
    Dim sigWidth As Double
    Dim sigHeight As Double
    Dim sigAspectRatio As Double
    Dim signatureImagePath As String
    
    ' These numbers make no sense, but they work.
    Const ABSOLUTE_LEFT As Double = 375
    Const ABSOLUTE_TOP As Double = 727.5
    Const MAX_WIDTH As Double = 130
    Const MAX_HEIGHT As Double = 31
    
    ' Step 1: Check if signature is already present in the report template
    On Error Resume Next
    Set sigShape = pptDoc.Slides(1).Shapes(sigShapeName)
    
    signatureImagePath = ExportSignatureFromExcel(sigShapeName)
    Set sigShape = pptDoc.Slides(1).Shapes.AddPicture(fileName:=signatureImagePath, _
                                                      LinkToFile:=msoFalse, _
                                                      SaveWithDocument:=msoTrue, _
                                                      Left:=ABSOLUTE_LEFT, _
                                                      Top:=ABSOLUTE_TOP)
    
    sigShape.Name = sigShapeName
    
    If Err.Number <> 0 Then
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.Reports.ErrorInsertingSignature")
        End If
        Exit Sub
    End If
    On Error GoTo 0
    
    sigAspectRatio = sigShape.Width / sigShape.Height   ' Maintain the aspect ratio and resize if needed
    If MAX_WIDTH / MAX_HEIGHT > sigAspectRatio Then
        sigWidth = MAX_HEIGHT * sigAspectRatio          ' Adjust width to fit within max height
        sigHeight = MAX_HEIGHT
    Else
        sigWidth = MAX_WIDTH                            ' Adjust height to fit within max width
        sigHeight = MAX_WIDTH / sigAspectRatio
    End If

    With sigShape
        .LockAspectRatio = msoTrue
        .Width = sigWidth
        .Height = sigHeight
    End With
End Sub

Private Function ExportSignatureFromExcel(ByVal sigShapeName As String) As String
    Dim tempSheet           As Worksheet
    Dim signatureshp        As Shape
    Dim chrtObj             As ChartObject
    Dim signatureImagePath  As String
    
    Application.DisplayAlerts = False
    
    Set tempSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    tempSheet.Name = "Temp_signature"
    
    Set signatureshp = Options.Shapes(sigShapeName)
    signatureshp.Copy
    
    signatureImagePath = GetDefaultFolderPaths("Temp")
    If Not CheckForAndAttemptToCreateFolder(signatureImagePath) Then
        signatureImagePath = GetDefaultFolderPaths("Resources")
    End If
    
    signatureImagePath = signatureImagePath & "tempSignature.png"
        
    On Error Resume Next
    Kill signatureImagePath
    Err.Clear
    
    Set chrtObj = tempSheet.ChartObjects.Add(Left:=tempSheet.Range("B2").Left, _
                                             Top:=tempSheet.Range("B2").Top, _
                                             Width:=signatureshp.Width, _
                                             Height:=signatureshp.Height)
    With chrtObj
        .Activate
        With .chart
            .Paste
            .ChartArea.Format.Line.Visible = msoFalse
            .Export signatureImagePath, "png"
        End With
        .Delete
    End With
    On Error GoTo 0
    
    tempSheet.Delete
    Application.DisplayAlerts = True
    
    ExportSignatureFromExcel = signatureImagePath
End Function

Private Function GetSignatureFile(ByVal signaturePath As String) As String
    #If Mac Then
        On Error Resume Next
        GetSignatureFile = AppleScriptTask(APPLE_SCRIPT_FILE, "FindSignature", signaturePath)
        On Error GoTo 0
    #Else
        If DoesFileExist(signaturePath & "mySignature.png") Then
            GetSignatureFile = signaturePath & "mySignature.png"
        ElseIf DoesFileExist(signaturePath & "mySignature.jpg") Then
            GetSignatureFile = signaturePath & "mySignature.jpg"
        Else
            GetSignatureFile = vbNullString
        End If
    #End If
End Function

Private Function WriteFileToDisk(ByRef ws As Worksheet, ByRef pptApp As Object, ByRef pptDoc As Object, ByVal generateProcess As String, ByVal fileName As String) As Boolean
    Const viewNotesMaster As Long = 5
    Const viewNormal      As Long = 9
    
    Dim tempFile As String
    Dim destFile As String
    Dim savePath As String
    Dim subfolderPath As String
    
    #If Mac Then
        Dim scriptResult As Boolean
    #Else
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
    #End If
    
    savePath = SetSavePath(ws, generateProcess)
    If savePath = vbNullString Then
        ' Set an error msg
        WriteFileToDisk = False
        Exit Function
    End If
    
    fileName = fileName & IIf(generateProcess = "Proofs", ".pptx", ".pdf")
    tempFile = GetDefaultFolderPaths("Temp") & fileName
    destFile = savePath & fileName
    
    On Error Resume Next
    DeleteFile tempFile
    
    SwitchPptViewType pptApp, viewNormal
    
    Select Case generateProcess
        Case "Proofs"
            pptDoc.SaveCopyAs tempFile
        Case "Reports", "Certificates"
            #If Mac Then
                scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "SavePptAsPdf", tempFile)
            #Else
                pptDoc.ExportAsFixedFormat Path:=tempFile, FixedFormatType:=2, Intent:=2, PrintRange:=Nothing, BitmapMissingFonts:=True
            #End If
    End Select

    SwitchPptViewType pptApp, viewNotesMaster
    
    #If Mac Then
        scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CopyFile", tempFile & APPLE_SCRIPT_SPLIT_KEY & destFile)
    #Else
        If fso.fileExists(tempFile) Then fso.CopyFile tempFile, destFile, True
    #End If
    On Error GoTo 0
    
    If g_UserOptions.EnableLogging Then
        If Err.Number = 0 Then
            DebugAndLogging GetMsg("Debug.Reports.Saved")
        Else
            DebugAndLogging GetMsg("Debug.Reports.ErrorSaving", Err.Number, Err.Description)
        End If
    End If
    
    DisableAutoSave pptApp, pptDoc
    
    WriteFileToDisk = (Err.Number = 0)
End Function

Private Sub SwitchPptViewType(ByRef pptApp As Object, ByVal viewType As Long)
    pptApp.ActiveWindow.viewType = viewType
End Sub

Private Sub OpenDestinationFolder(ByVal savePath As Variant)
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.FileManagement.OpenDestinationFolder", savePath)
    End If
    
#If Mac Then
    Dim scriptResult As Boolean
    
    On Error Resume Next
    scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "OpenFolder", savePath)
    On Error GoTo 0
#Else
    Shell "explorer.exe """ & savePath & """", vbNormalFocus
#End If
End Sub

Private Sub DisableAutoSave(ByRef pptApp As Object, ByRef pptDoc As Object)
    Dim isDisabled As Boolean
    
    If Val(pptApp.Version) <= 15 Then
        Exit Sub
    End If

    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.PowerPoint.DisableAutoSave")
    End If

    On Error Resume Next
    pptDoc.AutoSaveOn = False
    isDisabled = (pptDoc.AutoSaveOn = False)
    On Error GoTo 0

    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.CodeExecution.Result", INDENT_LEVEL_1, IIf(isDisabled, "Enabled", "Disabled"))
    End If
End Sub

Private Sub SetPowerPointViewSettings(ByRef pptApp As Object)
    With pptApp
        On Error Resume Next
        .ActiveWindow.viewType = 5 ' SliderSorterView to reduce redraw load
        .WindowState = 2 ' Minimized
        On Error GoTo 0
    End With
End Sub

Private Function GetElementPositioning(ByVal CertificateDesign As String) As String
    Select Case CertificateDesign
        Case "Default", "Modern"
            GetElementPositioning = "Standard"
    End Select
End Function

Private Sub ToggleCertificateStyleShapes(ByRef pptDoc As Object, ByVal borderType As String, ByVal borderColorCode As String, ByVal elementPositioning As String)
    Dim currentGrp As Object
    Dim currentShp As Shape
    Dim shpName As String
    Dim shpArray As Variant
    Dim i As Long
    Dim j As Long

    shpArray = Array("Base_Elements", "Modified_Elements", "Trophies", "Emblems", "Placements", "Levels")
    
    With pptDoc.Slides(1).Shapes
        For i = LBound(shpArray) To UBound(shpArray)
            Set currentGrp = .Item(shpArray(i))
            
            currentGrp.Visible = msoTrue

            With currentGrp.GroupItems
                Select Case shpArray(i)
                    Case "Base_Elements", "Modified_Elements"
                        For j = .Count To 1 Step -1
                            .Item(j).Visible = (Right$(.Item(j).Name, Len(elementPositioning)) = elementPositioning)
                        Next j
                    ' For efficiency, hide all of these shapes here
                    Case "Trophies", "Emblems", "Placements", "Levels"
                        For j = .Count To 1 Step -1
                            .Item(j).Visible = msoFalse
                        Next j
                End Select
            End With
        Next i
        
        .Item("Borders").Visible = (borderType <> "Disabled")
        If borderType <> "Disabled" Then
            ToggleBorderVisibility .Item("Borders"), borderType, borderColorCode
        End If
    End With
End Sub

Private Sub ImportTextToCertificates(ByRef pptDoc As Object, ByVal elementPositioning As String, ByVal importStage As String, ByVal firstValue As String, ByVal secondValue As String, Optional ByVal thirdValue As String = vbNullString)
    Dim fullName As String
    Dim koreanTextLength As Long
    
    On Error Resume Next
    Select Case importStage
        Case "Stage1"
            With pptDoc.Slides(1).Shapes
                With .Item("Modified_Elements").GroupItems
                    .Item("Teacher_" & elementPositioning).TextFrame.TextRange.text = "in " & firstValue & " Teacher's"
                    .Item("Date_Text_" & elementPositioning).TextFrame.TextRange.text = thirdValue
                End With
                .Item("Levels").GroupItems(secondValue & "_" & elementPositioning).Visible = msoTrue
            End With
        Case "Stage2"
            fullName = secondValue
            If Len(firstValue) < 10 Then
                fullName = fullName & " (" & firstValue & ")"
            End If
            koreanTextLength = Len(secondValue)
            With pptDoc.Slides(1).Shapes("Modified_Elements").GroupItems("Student_" & elementPositioning).TextFrame.TextRange
                .text = fullName
                .Characters(1, koreanTextLength).Font.Name = "Kakao Big Sans"
                .Characters(koreanTextLength + 1, Len(fullName) - koreanTextLength).Font.Name = "Constantia"
            End With
    End Select
    On Error GoTo 0
End Sub

Private Sub ToggleCertificateShapesVisiblity(ByRef certificateShape As Object, Optional ByVal elementPositioning As String = vbNullString)
    Dim i As Long

    With certificateShape.GroupItems
        Select Case certificateShape.Name
            Case "Base_Elements", "Modified_Elements"
                For i = .Count To 1 Step -1
                    .Item(i).Visible = (Right$(.Item(i).Name, Len(elementPositioning)) = elementPositioning)
                Next i
            Case "Trophies", "Emblems"
                For i = .Count To 1 Step -1
                    If .Item(i).Visible Then .Item(i).Visible = msoFalse
                Next i
            Case "Placements", "Levels"
                For i = .Count To 1 Step -1
                    If .Item(i).Visible Then .Item(i).Visible = msoFalse
                Next i
        End Select
    End With
End Sub

Private Sub ToggleStudentPlacementShapes(ByRef pptDoc As Object, ByVal elementPositioning As String, ByVal CertificateDesign As String, ByVal studentRanking As Long)
    Dim trophyGroup     As Object
    Dim emblemsGroup    As Object
    Dim placementGroup  As Object
    Dim rankings        As Variant
    Dim placements      As Variant
    Dim rankingToShow   As String
    Dim isVisible       As Boolean
    Dim i               As Long
    
    rankings = Array("Gold_", "Silver_", "Bronze_")
    placements = Array("First_", "Second_", "Third_")
    
    With pptDoc.Slides(1).Shapes
        Set trophyGroup = .Item("Trophies").GroupItems
        Set emblemsGroup = .Item("Emblems").GroupItems
        Set placementGroup = .Item("Placements").GroupItems
    End With
    
    For i = LBound(rankings) To UBound(rankings)
        isVisible = (studentRanking = i + 1)
        rankingToShow = rankings(i) & CertificateDesign
        
        trophyGroup.Item(rankingToShow).Visible = isVisible
        emblemsGroup.Item(rankingToShow).Visible = isVisible
        placementGroup.Item(placements(i) & elementPositioning).Visible = isVisible
    Next i
End Sub

Private Sub ToggleBorderVisibility(ByRef borderShape As Object, ByVal borderStyle As String, ByVal borderColorCode As String)
    Dim i As Long
    
    With borderShape.GroupItems
        For i = .Count To 1 Step -1
            .Item(i).Visible = (.Item(i).Name = borderStyle)
            If .Item(i).Visible Then .Item(i).Fill.ForeColor.RGB = ConvertHexToRGB(borderColorCode)
        Next i
    End With
End Sub