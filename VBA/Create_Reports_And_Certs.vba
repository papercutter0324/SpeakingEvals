Option Explicit

#Const PRINT_DEBUG_MESSAGES = True
#If Mac Then
    Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
    Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
#End If

Public Sub CreateReportsAndCertificates(ByVal ws As Worksheet, ByVal clickedButtonName As String)
    Const CERTIFICATE_TEMPLATE As String = "CertificateTemplate.pptx"
    Const REPORT_TEMPLATE As String = "SpeakingEvaluationTemplate.pptx"
    
    Const ERR_RESOURCES_FOLDER As String = "resourcesFolder"
    Const ERR_FONT_INSTALLATION As String = "fontInstallation"
    Const ERR_INCOMPLETE_RECORDS As String = "incompleteRecords"
    Const ERR_LOADING_POWERPOINT As String = "loadingPowerPoint"
    Const ERR_LOADING_TEMPLATE As String = "loadingTemplate"
    Const ERR_MISSING_SHAPES As String = "missingTemplateShapes"
    Const MSG_SAVE_FAILED As String = "exportFailed"
    Const MSG_ZIP_FAILED As String = "zipFailed"
    Const MSG_SUCCESS As String = "exportSuccessful"
    
    ' PPT app and template objects
    Dim pptApp As Object
    Dim pptDoc As Object
    
    ' Key values loaded from the current class's records
    Dim winnerValues(1 To 10) As String
    Dim firstStudentRecord As Long
    Dim lastRow As Long
    Dim saveResult As Boolean
    
    ' Strings for user and debug messages
    Dim debugMsgValue As String
    Dim resultMsg As String
    Dim msgToDisplay As String
    Dim msgTitle As String
    Dim msgType As Long
    Dim dialogSize As Long
    Dim msgResult As Variant
    
    ' Important file and folder values
    Dim resourcesFolder As String
    Dim templateName As String
    Dim templatePath As String
    Dim templateToUse As String
    Dim savePath As String
    Dim subfolderPath As String
    
    Dim generateProcess As String
    
    #If Mac Then
        Dim scriptResult As Boolean
    #End If
    
    ' Step 1: Determine types of reports to generate
    Select Case clickedButtonName
        Case "Button_GenerateReports"
            generateProcess = "FinalReports"
            debugMsgValue = "Final Reports"
            templateName = REPORT_TEMPLATE
            subfolderPath = vbNullString
        Case "Button_GenerateProofs"
            generateProcess = "Proofs"
            debugMsgValue = "Report Proofs"
            templateName = REPORT_TEMPLATE
            subfolderPath = vbNullString
        Case "Button_GenerateCertificates"
            generateProcess = "Certificates"
            debugMsgValue = "Winner Certificates"
            templateName = CERTIFICATE_TEMPLATE
            subfolderPath = "Certificates"
        Case Else
            msgToDisplay = "You have clicked an invalid option for creating the reports. This shouldn't be possible unless this file has been altered " & _
                           "in an unintended manner. Please download a new copy of this Excel file, copy over all of the students' records, and try again."
            msgResult = DisplayMessage(msgToDisplay, vbExclamation, "Invalid Selection!")
        Exit Sub
    End Select
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Beginning File Creation" & vbNewLine & _
                    INDENT_LEVEL_1 & "File Types: " & debugMsgValue
    #End If
    
    ' Step 2: Ensure ./Resources exists
    resourcesFolder = ConvertOneDriveToLocalPath(ThisWorkbook.Path & Application.PathSeparator & "Resources")
    If Not CheckForFolder(resourcesFolder) Then
        resultMsg = ERR_RESOURCES_FOLDER
        GoTo CleanUp
    End If
    
    ' Step 3: Locate required fonts and install if missing
    If Not InstallFonts() Then
        resultMsg = ERR_FONT_INSTALLATION
        GoTo CleanUp
    End If
    
    ' Step 4: Check if template is open and close if necessary
    templatePath = resourcesFolder & Application.PathSeparator & templateName
    If IsPptTemplateAlreadyOpen(resourcesFolder, templateName) Then ' This can be updated to simply pass templatePath
        ' Set an error msg
        ' resultMsg = ERR_PPT_OPEN
        GoTo CleanUp
    End If

    ' Step 5: Locate valid copy of needed template
    templateToUse = LocateRequiredFile(templateName, templatePath, "Template")
    If templateToUse = vbNullString Then
        ' Set an error msg
        ' resultMsg = ERR_MISSING_TEMPLATE
        GoTo CleanUp
    End If
    
    ' Step 6: Verify required information is present and/or loaded
    Select Case generateProcess
        Case "FinalReports", "Proofs"
            If Not VerifyRecordsAreComplete(ws, lastRow, firstStudentRecord) Then
                resultMsg = ERR_INCOMPLETE_RECORDS
                GoTo CleanUp
            End If
        Case "Certificates"
            LoadWinnerValues ws, winnerValues()
    End Select
    
    ' Step 7: Set (and create if necessary) folder to save files to
    savePath = SetSavePath(ws, subfolderPath)
    If savePath = vbNullString Then
        ' Set an error msg
        ' resultMsg = ERR_INVALID_SAVE_PATH
        GoTo CleanUp
    End If
    
    ' Step 8: Grab an open instance of PowerPoint or open a new one
    If Not LoadPowerPoint(pptApp, pptDoc, templateToUse) Then
        resultMsg = ERR_LOADING_POWERPOINT
        GoTo CleanUp
    End If
    
    ' Step 8a: Early exit if there was a problem opening PowerPoint or the template
    If pptDoc Is Nothing Then
        resultMsg = ERR_LOADING_TEMPLATE
        GoTo CleanUp
    End If
    
    ' Step 9: Generate requested files
    Select Case generateProcess
        Case "FinalReports", "Proofs"
            saveResult = CreateReportFiles(ws, generateProcess, pptApp, pptDoc, lastRow, firstStudentRecord, savePath)
        Case "Certificates"
            saveResult = CreateCertificateFiles(ws, pptApp, pptDoc, savePath, winnerValues())
    End Select
    
    ' Step 9a: Early exit if an error was encountered while saving reports
    If Not saveResult Then
        resultMsg = MSG_SAVE_FAILED
        GoTo CleanUp
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print INDENT_LEVEL_1 & "Save process complete."
    #End If
    
    ' Step 10: Close the open instance of the template and PowerPoint
    KillPowerPoint pptApp, pptDoc
    resultMsg = MSG_SUCCESS
    
    ' Step 11: Generatea zip file if "Generate Reports" was selected
    If generateProcess = "FinalReports" Then
        ZipReports ws, savePath, saveResult, resourcesFolder
        If Not saveResult Then resultMsg = MSG_ZIP_FAILED
    End If
    
    ' Step 12: Open folder where files were saved
    If saveResult Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Attempting to open destination folder." & vbNewLine & _
                        INDENT_LEVEL_1 & "Path: " & savePath
        #End If
        
        ' MacOS: Opening the folder is not currently supported, so inform the user of the path
        #If Mac Then
            msgToDisplay = "Generated reports have been saved to: " & vbNewLine & savePath
            msgTitle = "Notice!"
            msgType = vbInformation
            dialogSize = 350
            msgResult = DisplayMessage(msgToDisplay, msgType, msgTitle, dialogSize)
        #Else
            Shell "explorer.exe """ & savePath & """", vbNormalFocus
        #End If
    End If
CleanUp:
    ' Step 13: Set appropriate message to display to the user
    Select Case resultMsg
        Case ERR_RESOURCES_FOLDER
            msgToDisplay = "Unable to locate or create the Resources folder. Please create this manually in the same location as this spreadsheet and try again."
            msgTitle = "Error!"
            msgType = vbExclamation
            dialogSize = 330
        Case ERR_INCOMPLETE_RECORDS
            msgToDisplay = "One or more fields for missing. Please complete all fields and try again."
            msgTitle = "Missing Data!"
            msgType = vbExclamation
            dialogSize = 230
        Case ERR_FONT_INSTALLATION
            msgToDisplay = "There was an error when trying to install the required font. Please try again, and if the error persists, consider installing" & vbNewLine & vbNewLine & _
                           "the font manually. You can find the link on the Instructions sheet."
            msgTitle = "Font Error!"
            msgType = vbExclamation
            dialogSize = 360
        Case ERR_LOADING_POWERPOINT, ERR_LOADING_TEMPLATE
            msgToDisplay = "Error opening PowerPoint and/or the template." & vbNewLine & vbNewLine & "PowerPoint may have encountered a bug preventing it from opening " & _
                           "and/or closing properly. Please wait a couple seconds and try again. If this error persists, please try rebooting your computer."
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
    
    If resultMsg <> vbNullString Then msgResult = DisplayMessage(msgToDisplay, msgType, msgTitle, dialogSize)
    
    ' Step 14: Final check to close the template and PowerPoint in case of an early exit
    If Not pptApp Is Nothing Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Beginning final cleanup checks."
        #End If
        KillPowerPoint pptApp, pptDoc
    End If
End Sub

Private Function IsPptTemplateAlreadyOpen(ByVal resourcesFolder As String, ByVal templateName As String) As Boolean
    Dim pptApp As Object
    Dim pptDoc As Object
    Dim templatePath As String
    Dim templateIsOpen As Boolean
    Dim pathOfOpenDoc As String
    Dim msgToDisplay As String
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Checking for Existing PowerPoint Instances"
    #End If
    
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    Err.Clear
    
    If Not pptApp Is Nothing Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print INDENT_LEVEL_1 & "PowerPoint instance found" & vbNewLine & _
                        INDENT_LEVEL_1 & "Checking if template is already open "
        #End If
        
        templatePath = resourcesFolder & Application.PathSeparator & templateName
        
        For Each pptDoc In pptApp.Presentations
            pathOfOpenDoc = ConvertOneDriveToLocalPath(pptDoc.fullName)
            If StrComp(pathOfOpenDoc, templatePath, vbTextCompare) = 0 Then
                templateIsOpen = True
                
                #If PRINT_DEBUG_MESSAGES Then
                    Debug.Print INDENT_LEVEL_1 & "Open instance found" & vbNewLine & _
                                INDENT_LEVEL_1 & "Asking if user wishes to automatically close and continue."
                #End If
                
                msgToDisplay = "An open instance of MS PowerPoint has been detected. Please save any open files before continuing." & vbNewLine & vbNewLine & _
                               "Click OK to automatically close PowerPoint and continue, or click Cancel to finish and save your work."
                               
                If DisplayMessage(msgToDisplay, vbOKCancel + vbCritical, "Notice!", 310) = vbOK Then
                    pptDoc.Close SaveChanges:=False
                    templateIsOpen = False
                    #If PRINT_DEBUG_MESSAGES Then
                        Debug.Print INDENT_LEVEL_1 & "Open instance has been closed."
                    #End If
                End If
            End If
        Next pptDoc
    End If
    On Error GoTo 0
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print INDENT_LEVEL_1 & "Open instances: " & templateIsOpen
    #End If
    
    Set pptDoc = Nothing
    Set pptApp = Nothing
    IsPptTemplateAlreadyOpen = templateIsOpen
End Function

Private Sub LoadWinnerValues(ByVal ws As Worksheet, ByRef winnerValues() As String)
    Dim topThreeStudents As Variant
    Dim winnerNameHalves() As String
    Dim tempValue As String
    Dim winnerIndex As Long
    Dim i As Long
    
    ' Step 1: Load values for Native Teacher, Level, Class Days, and the winning students
    With ws
        winnerValues(1) = Trim$(.Range("C1"))
        winnerValues(2) = Trim$(.Range("C3"))
        winnerValues(3) = Trim$(.Range("C4"))
        winnerValues(4) = Trim$(.Range("C6"))
        topThreeStudents = .Range("L2:L4").Value
    End With
    
    ' Step 2: Split the winners' English and Korean names for efficient use later
    winnerIndex = 5 ' Starting index for the winners
    For i = 1 To 3
        If topThreeStudents(i, 1) <> vbNullString And topThreeStudents(i, 1) <> "Incomplete List" Then
            tempValue = Left$(topThreeStudents(i, 1), Len(topThreeStudents(i, 1)) - 1)
            winnerNameHalves() = Split(tempValue, "(")
            
            If UBound(winnerNameHalves) = 1 Then
                winnerValues(winnerIndex) = Trim$(winnerNameHalves(0))
                winnerValues(winnerIndex + 1) = Trim$(winnerNameHalves(1))
            Else
                winnerValues(winnerIndex) = vbNullString
                winnerValues(winnerIndex + 1) = vbNullString
            End If
        Else
            winnerValues(winnerIndex) = vbNullString
            winnerValues(winnerIndex + 1) = vbNullString
        End If
        
        ' Step 2a: Iterate by two for the next winner's index
        winnerIndex = winnerIndex + 2
    Next i
End Sub

Private Function LoadPowerPoint(ByRef pptApp As Object, ByRef pptDoc As Object, ByVal templatePath As String) As Boolean
    #If Mac Then
        Dim appleScriptResult As String
        Dim msgToDisplay As String
        Dim msgResult As Variant
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Opening PowerPoint"
    #End If
    
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    Err.Clear
    On Error GoTo ErrorHandler
    
    ' Open a new instance of PowerPoint if needed
    #If Mac Then
        If pptApp Is Nothing Then
            appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "LoadApplication", "Microsoft PowerPoint")
            
            #If PRINT_DEBUG_MESSAGES Then
                If appleScriptResult <> "" Then Debug.Print appleScriptResult
            #End If
            
            appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "IsAppLoaded", "Microsoft PowerPoint")
            
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print INDENT_LEVEL_1 & "" & appleScriptResult
            #End If
            
            Set pptApp = GetObject(, "PowerPoint.Application")
        End If
    #Else
        If pptApp Is Nothing Then Set pptApp = CreateObject("PowerPoint.Application")
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print INDENT_LEVEL_1 & "PowerPoint Loaded: " & (Not pptApp Is Nothing)
    #End If
    
    ' Make the process visible so users understand their computer isn't frozen
    pptApp.Visible = True
    
    If Not pptApp Is Nothing Then
        Set pptDoc = pptApp.Presentations.Open(templatePath)
        If Val(pptApp.Version) > 15 Then
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print INDENT_LEVEL_1 & "Disabling AutoSave"
            #End If
            DisableAutoSave pptDoc
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print INDENT_LEVEL_1 & "Status: " & pptDoc.AutoSaveOn
            #End If
        End If
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print INDENT_LEVEL_1 & "Template loaded: " & (Not pptDoc Is Nothing)
    #End If
    
    LoadPowerPoint = (Not pptApp Is Nothing)
    Exit Function
ErrorHandler:
    #If Mac Then
        msgToDisplay = "An error occurred while trying to load Microsoft PowerPoint. This is usually a result of a quirk in MacOS. Try creating the reports again, and it should work fine." & vbNewLine & vbNewLine & _
                        "If the problem persists, please take a picture of the following error message and ask your team leader to send it to Warren at Bundang." & vbNewLine & vbNewLine & _
                        "VBA Error " & Err.Number & ": " & Err.Description & vbNewLine & "AppleScript Error: " & appleScriptResult
        msgResult = DisplayMessage(msgToDisplay, vbOKOnly, "Error Loading PowerPoint", 470)
    #End If
    LoadPowerPoint = False
End Function

Private Function CreateReportFiles(ByVal ws As Worksheet, ByVal generateProcess As String, ByVal pptApp As Object, ByVal pptDoc As Object, ByVal lastRow As Long, ByVal firstStudentRecord As Long, ByVal savePath As String) As Boolean
    Dim saveResult As Boolean
    Dim currentRow As Long
    Dim i As Long
    
    For currentRow = firstStudentRecord To lastRow
        #If PRINT_DEBUG_MESSAGES Then
            i = i + 1
            Debug.Print INDENT_LEVEL_1 & "Generating Report " & i & " of " & (lastRow - firstStudentRecord + 1)
        #End If
        WritePptReport ws, pptApp, pptDoc, generateProcess, currentRow, savePath, saveResult
        
        If Not saveResult Then Exit For
    Next currentRow
    
    CreateReportFiles = saveResult
End Function

Private Function CreateCertificateFiles(ByVal ws As Worksheet, ByVal pptApp As Object, ByVal pptDoc As Object, ByVal savePath As String, ByRef winnerValues() As String) As Boolean
    Dim enableBorder As Boolean
    Dim certificateCreated As Boolean
    Dim fileName As String
    Dim certificateSettings As Variant
    Dim studentRanking As Long
    Dim i As Long
    
    
    ' Step 1: Check Winner Certificates sheet for layout
    '   WIP Options: Borderless (Default), Bordered, Default, Modern (aka flat), etc
    '   To be added elsewhere, but display a preview of selected style for the user
    certificateSettings = Options.Range("K10:K14").Value
    
    ' Step 2: Toggle visibility of design shapes for chosen style
    ToggleCertificateStyleShapes pptDoc, certificateSettings
    
    ' Step 3: Import the teacher's name, class's level, and date
    ImportTextToCertificates pptDoc, certificateSettings(2, 1), "Stage1", winnerValues(1), winnerValues(2), winnerValues(4)
    
    ' Step 4: Create certificates
    studentRanking = 0
    For i = 5 To 9 Step 2 ' Step indexes for each student in winnerValues()
        studentRanking = studentRanking + 1
        If winnerValues(i) <> vbNullString Then
            ' Step 4a: Display correct ribbon and trophy
            ToggleStudentPlacementShapes pptDoc, certificateSettings(2, 1), studentRanking
            
            ' Step 4b: Import the current student's name
            ImportTextToCertificates pptDoc, certificateSettings(2, 1), "Stage2", winnerValues(i), winnerValues(i + 1)
    
            ' Step 5: Save certificate as a PDF
            fileName = winnerValues(i + 1) & "(" & winnerValues(i) & ") - " & winnerValues(3)
            certificateCreated = WriteFileToDisk(pptApp, pptDoc, "Certificates", savePath, fileName)
    
            ' Step 5a: Early exit if a write error occurs
            If Not certificateCreated Then
                'Display error message
                CreateCertificateFiles = False
                Exit Function
            End If
        End If
    Next i

    CreateCertificateFiles = True
End Function

Private Sub ToggleCertificateStyleShapes(ByVal pptDoc As Object, ByVal certificateSettings As Variant)
    Dim certificateLayout As String
    Dim certificateDesign As String
    Dim borderStyle As String
    Dim borderColorCode As String
    Dim i As Long
    
    certificateLayout = certificateSettings(1, 1)
    certificateDesign = certificateSettings(2, 1)
    borderStyle = certificateSettings(3, 1)
    borderColorCode = certificateSettings(5, 1)
    
    ' Step 1: Toggle correct page layout
    With pptDoc.Slides(1).Shapes
        If Not .Item("Base_Elements").Visible Then .Item("Base_Elements").Visible = msoTrue
        ToggleCertificateShapesVisiblity .Item("Base_Elements"), certificateDesign
        
        If Not .Item("Modified_Elements").Visible Then .Item("Modified_Elements").Visible = msoTrue
        ToggleCertificateShapesVisiblity .Item("Modified_Elements"), certificateDesign
        
        If Not .Item("Trophies").Visible Then .Item("Trophies").Visible = msoTrue
        ToggleCertificateShapesVisiblity .Item("Trophies")
        
        If Not .Item("Emblems").Visible Then .Item("Emblems").Visible = msoTrue
        ToggleCertificateShapesVisiblity .Item("Emblems")
        
        If Not .Item("Placements").Visible Then .Item("Placements").Visible = msoTrue
        ToggleCertificateShapesVisiblity .Item("Placements")
        
        If Not .Item("Levels").Visible Then .Item("Levels").Visible = msoTrue
        ToggleCertificateShapesVisiblity .Item("Levels")
        
        .Item("Borders").Visible = (borderStyle <> "Disabled")
        If borderStyle <> "Disabled" Then ToggleBorderVisibility .Item("Borders"), borderStyle, borderColorCode
    End With
End Sub

Private Sub ImportTextToCertificates(ByVal pptDoc As Object, ByVal certificateStyle As String, ByVal importStage As String, ByVal firstValue As String, ByVal secondValue As String, Optional ByVal thirdValue As String = vbNullString)
    Dim targetGroup As Object
    Dim fullName As String
    Dim koreanTextLength As Long
    
    On Error Resume Next
    Select Case importStage
        Case "Stage1"
            With pptDoc.Slides(1).Shapes
                With .Item("Modified_Elements").GroupItems
                    .Item("Teacher_" & certificateStyle).TextFrame.TextRange.Text = "in " & firstValue & " Teacher's"
                    .Item("DateText_" & certificateStyle).TextFrame.TextRange.Text = thirdValue
                End With
                .Item("Levels").GroupItems(secondValue & "_" & certificateStyle).Visible = msoTrue
            End With
        Case "Stage2"
            fullName = secondValue
            If Len(firstValue) < 10 Then
                fullName = fullName & " (" & firstValue & ")"
            End If
            koreanTextLength = Len(secondValue)
            With pptDoc.Slides(1).Shapes("Modified_Elements").GroupItems("Student_" & certificateStyle).TextFrame.TextRange
                .Text = fullName
                .Characters(1, koreanTextLength).Font.Name = "Kakao Big Sans"
                .Characters(koreanTextLength + 1, Len(fullName) - koreanTextLength).Font.Name = "Constantia"
            End With
    End Select
    On Error GoTo 0
End Sub

Private Sub ToggleCertificateShapesVisiblity(ByVal certificateShape As Object, Optional ByVal certificateDesign As String = vbNullString)
    Dim i As Long

    With certificateShape.GroupItems
        Select Case certificateShape.Name
            Case "Base_Elements", "Modified_Elements"
                For i = .Count To 1 Step -1
                    .Item(i).Visible = (Right$(.Item(i).Name, Len(certificateDesign)) = certificateDesign)
                Next i
            Case "Trophies", "Emblems", "Placements", "Levels"
                For i = .Count To 1 Step -1
                    If .Item(i).Visible Then .Item(i).Visible = msoFalse
                Next i
        End Select
    End With
End Sub

Private Sub ToggleBorderVisibility(ByVal borderShape As Object, ByVal borderStyle As String, ByVal borderColorCode As String)
    Dim i As Long
    
    With borderShape.GroupItems
        For i = .Count To 1 Step -1
            .Item(i).Visible = (.Item(i).Name = borderStyle)
            If .Item(i).Visible Then .Item(i).Fill.ForeColor.RGB = ConvertHexToRGB(borderColorCode)
        Next i
    End With
End Sub

Private Sub ToggleStudentPlacementShapes(ByVal pptDoc As Object, ByVal certificateStyle As String, ByVal studentRanking As Long)
    Dim trophyGroup As Object
    Dim emblemsGroup As Object
    Dim placementGroup As Object
    Dim rankings As Variant
    Dim placements As Variant
    Dim rankingToShow As String
    Dim isVisible As Boolean
    Dim i As Long
    
    rankings = Array("Gold_", "Silver_", "Bronze_")
    placements = Array("First_", "Second_", "Third_")
    
    With pptDoc.Slides(1).Shapes
        Set trophyGroup = .Item("Trophies").GroupItems
        Set emblemsGroup = .Item("Emblems").GroupItems
        Set placementGroup = .Item("Placements").GroupItems
    End With
    
    For i = LBound(rankings) To UBound(rankings)
        isVisible = (studentRanking = i + 1)
        rankingToShow = rankings(i) & certificateStyle
        
        trophyGroup.Item(rankingToShow).Visible = isVisible
        emblemsGroup.Item(rankingToShow).Visible = isVisible
        placementGroup.Item(placements(i) & certificateStyle).Visible = isVisible
    Next i
End Sub

Private Sub WritePptReport(ByVal ws As Object, ByVal pptApp As Object, ByVal pptDoc As Object, ByVal generateProcess As String, ByVal currentRow As Long, ByVal savePath As String, ByRef saveResult As Boolean)
    Dim reportMetaData As Variant
    Dim currentStudentData As Variant
    Dim englishName As String
    Dim koreanName As String
    Dim classLevel As String
    Dim nativeTeacher As String
    Dim koreanTeacher As String
    Dim evalDate As String
    Dim commentText As String
    Dim fileName As String
    Dim validEnglishName As String
    Dim classDay As String
    Dim classTime As String
    Dim scoreCategories As Variant
    Dim scoreValues As Variant
    Dim longEnglishName As Boolean
    Dim englishNameTextboxHeight As Long
    Dim englishNameTextboxTop As Long
    Dim englishNameFontSize As Long
    Dim i As Long
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print INDENT_LEVEL_2 & "Preparing Report Data"
    #End If
    
    ' Step 1: Bulk read
    With ws
        reportMetaData = .Range("C1:C6").Value
        currentStudentData = .Range("B" & currentRow & ":J" & currentRow).Value
    End With
    
    ' Step 2: Map arrays to the report header
    ' MetaData indices are (row, 1) because it's a single column read
    nativeTeacher = reportMetaData(1, 1)
    koreanTeacher = reportMetaData(2, 1)
    classLevel = reportMetaData(3, 1)
    classDay = reportMetaData(4, 1)
    classTime = reportMetaData(5, 1)
    evalDate = Format$(CDate(reportMetaData(6, 1)), "DD MMM. YYYY")
    
    ' RowData indices are (1, column) because it's a single row read
    englishName = currentStudentData(1, 1)
    koreanName = currentStudentData(1, 2)
    commentText = currentStudentData(1, 9)
    scoreValues = Array(currentStudentData(1, 3), currentStudentData(1, 4), currentStudentData(1, 5), _
                        currentStudentData(1, 6), currentStudentData(1, 7), currentStudentData(1, 8), _
                        CalculateOverallGrade(ws, currentRow))
    
    ' Step 3: Prepare other required arrays and values
    scoreCategories = Array("Grammar_", "Pronunciation_", "Fluency_", "Manner_", "Content_", "Effort_", "Result_")
    ReformatEnglishName englishName, englishNameTextboxHeight, englishNameTextboxTop, englishNameFontSize
    validEnglishName = SanitizeFileName(englishName)
    fileName = koreanName & "(" & validEnglishName & ")" & " - " & classDay
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print INDENT_LEVEL_2 & "Filename: " & fileName
    #End If
    
    ' Step 4: Write to the report template
    With pptDoc.Slides(1).Shapes
        With .Item("Report_Header").GroupItems
            With .Item("English_Name")
                .Height = englishNameTextboxHeight
                .Top = englishNameTextboxTop
                With .TextFrame.TextRange
                    .Text = englishName
                    .Font.Size = englishNameFontSize
                End With
            End With
            .Item("Korean_Name").TextFrame.TextRange.Text = koreanName
            .Item("Grade_Level").TextFrame.TextRange.Text = classLevel
            .Item("Native_Teacher").TextFrame.TextRange.Text = nativeTeacher
            .Item("Korean_Teacher").TextFrame.TextRange.Text = koreanTeacher
            .Item("Eval_Date").TextFrame.TextRange.Text = evalDate
        End With
        
        
        For i = LBound(scoreCategories) To UBound(scoreCategories)
            ToggleScoreVisibility pptDoc, scoreCategories(i), scoreValues(i)
        Next i

        .Item("Comments").TextFrame.TextRange.Text = commentText

        ' Step 5: Insert signature if not alreay present
        On Error Resume Next
        If .Item("mySignature") Is Nothing Then InsertSignature pptDoc
        On Error GoTo 0
    End With
    
    ' Step 6: Attempt to save the report
    saveResult = WriteFileToDisk(pptApp, pptDoc, generateProcess, savePath, fileName)
End Sub

Private Sub ReformatEnglishName(ByRef englishName As String, ByRef englishNameTextboxHeight As Long, ByRef englishNameTextboxTop As Long, ByRef englishNameFontSize As Long)
    Dim spacePositions As Collection
    Dim textLength As Long
    Dim lowerBound As Long
    Dim upperBound As Long
    Dim posToReplace As Long
    Dim pos As Long
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

Private Function CalculateOverallGrade(ByVal ws As Worksheet, ByVal currentRow As Long) As String
    Dim scoreRangeValues As Variant
    Dim totalScore As Double
    Dim avgScore As Double
    Dim numericScore As Long
    Dim i As Long
    
    totalScore = 0
    
    ' Step 1: Load current student's scores into an array
    scoreRangeValues = ws.Range("D" & currentRow & ":I" & currentRow).Value
    
    ' Step 2: Tally up the student's scores
    For i = LBound(scoreRangeValues, 2) To UBound(scoreRangeValues, 2)
        Select Case CStr(scoreRangeValues(1, i))
            Case "A+": numericScore = 5
            Case "A": numericScore = 4
            Case "B+": numericScore = 3
            Case "B": numericScore = 2
            Case "C": numericScore = 1
        End Select
        totalScore = totalScore + numericScore
    Next i
    
    ' Step 3: Calculate the students Overall Score
    ' Custom Rounding: They're young, so let's be a little generous.
    avgScore = totalScore / 6
    
    If avgScore - Int(avgScore) >= 0.4 Then
        avgScore = Int(avgScore) + 1
    Else
        avgScore = Int(avgScore)
    End If
    
    Select Case avgScore
        Case 5: CalculateOverallGrade = "A+"
        Case 4: CalculateOverallGrade = "A"
        Case 3: CalculateOverallGrade = "B+"
        Case 2: CalculateOverallGrade = "B"
        Case 1: CalculateOverallGrade = "C"
    End Select
End Function

Private Sub ToggleScoreVisibility(ByVal pptDoc As Object, ByVal scoreCategory As String, ByVal scoreValue As String)
    With pptDoc.Slides(1).Shapes(scoreCategory & "Scores").GroupItems
        .Item(scoreCategory & "A+").Visible = IIf(scoreValue = "A+", msoTrue, msoFalse)
        .Item(scoreCategory & "A").Visible = IIf(scoreValue = "A", msoTrue, msoFalse)
        .Item(scoreCategory & "B+").Visible = IIf(scoreValue = "B+", msoTrue, msoFalse)
        .Item(scoreCategory & "B").Visible = IIf(scoreValue = "B", msoTrue, msoFalse)
        .Item(scoreCategory & "C").Visible = IIf(scoreValue = "C", msoTrue, msoFalse)
    End With
End Sub

Private Sub InsertSignature(ByVal pptDoc As Object)
    Dim sigShape As Object
    Dim sigWidth As Double
    Dim sigHeight As Double
    Dim sigAspectRatio As Double
    Dim signaturePath As String
    Dim signatureImagePath As String
    Dim useEmbeddedSignature As Boolean
    
    Const SIGNATURE_SHAPE_NAME As String = "mySignature"
    ' These numbers make no sense, but they work.
    Const ABSOLUTE_LEFT As Double = 375
    Const ABSOLUTE_TOP As Double = 727.5
    Const MAX_WIDTH As Double = 130
    Const MAX_HEIGHT As Double = 31
    
    ' Step 1: Check if signature is already present in the report template
    On Error Resume Next
    Set sigShape = pptDoc.Slides(1).Shapes(SIGNATURE_SHAPE_NAME)
    If Not sigShape Is Nothing Then
        ' Step 1a: If found, exit this sub early
        Exit Sub
    End If
    
    ' Step 2: Determine if an embedded or external signature will be used
    signaturePath = ConvertOneDriveToLocalPath(ThisWorkbook.Path & Application.PathSeparator)
    useEmbeddedSignature = (Not Options.Shapes.[_Default](SIGNATURE_SHAPE_NAME) Is Nothing)
    Err.Clear ' Clear the error if an embedded signature isn't found
     
    If useEmbeddedSignature Then
        ExportSignatureFromExcel SIGNATURE_SHAPE_NAME, signatureImagePath
    Else
        'Step 2a: If neither an embedded or external signature is found, exit early
        signatureImagePath = GetSignatureFile(signaturePath)
        If signatureImagePath = vbNullString Then Exit Sub
    End If
    
    ' Step 3: Add the signature into the report and name it "mySignature"
    Set sigShape = pptDoc.Slides(1).Shapes.AddPicture(fileName:=signatureImagePath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                                                      Left:=ABSOLUTE_LEFT, Top:=ABSOLUTE_TOP)
    sigShape.Name = SIGNATURE_SHAPE_NAME
    
    If Err.Number <> 0 Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Error inserting signature."
        #End If
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Step 4: Resize the signature to fit on the template
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

Private Sub ExportSignatureFromExcel(ByVal SIGNATURE_SHAPE_NAME As String, ByRef signatureImagePath As String)
    Dim tempSheet As Worksheet
    Dim signatureshp As Shape
    Dim chrtObj As ChartObject
    
    Application.DisplayAlerts = False
    
    Set tempSheet = ThisWorkbook.Sheets.Add(After:=Sheets.[_Default](Sheets.Count))
    tempSheet.Name = "Temp_signature"
    
    Set signatureshp = Options.Shapes.[_Default](SIGNATURE_SHAPE_NAME)
    signatureshp.Copy
    
    signatureImagePath = ConvertOneDriveToLocalPath(GetTempFilePath("tempSignature.png"))
    
    On Error Resume Next
    Kill signatureImagePath
    Err.Clear
    
    Set chrtObj = tempSheet.ChartObjects.Add(Left:=tempSheet.Range("B2").Left, Top:=tempSheet.Range("B2").Top, _
                                             Width:=signatureshp.Width, Height:=signatureshp.Height)
    With chrtObj
        .Activate
        With .Chart
            .Paste
            .ChartArea.Format.Line.Visible = msoFalse
            .Export signatureImagePath, "png"
        End With
        .Delete
    End With
    On Error GoTo 0
    
    tempSheet.Delete
    Application.DisplayAlerts = True
End Sub

Private Function GetSignatureFile(ByVal signaturePath As String) As String
    #If Mac Then
        GetSignatureFile = AppleScriptTask(APPLE_SCRIPT_FILE, "FindSignature", signaturePath)
    #Else
        If Dir(signaturePath & "mySignature.png") <> vbNullString Then
            GetSignatureFile = signaturePath & "mySignature.png"
        ElseIf Dir(signaturePath & "mySignature.jpg") <> vbNullString Then
            GetSignatureFile = signaturePath & "mySignature.jpg"
        Else
            GetSignatureFile = vbNullString
        End If
    #End If
End Function

Private Function WriteFileToDisk(ByVal pptApp As Object, ByVal pptDoc As Object, ByVal saveRoutine As String, ByVal savePath As String, ByVal fileName As String) As Boolean
    Dim tempFile As String
    Dim destFile As String
    
    #If Mac Then
        Dim scriptResult As Boolean
    #Else
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
    #End If
    
    fileName = fileName & IIf(saveRoutine = "Proofs", ".pptx", ".pdf")
    tempFile = GetTempFilePath(fileName)
    destFile = savePath & fileName
    
    On Error Resume Next
    DeleteFile tempFile
    
    Select Case saveRoutine
        Case "Proofs"
            pptDoc.SaveCopyAs tempFile
        Case "FinalReports", "Certificates"
            #If Mac Then
                scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "SavePptAsPdf", tempFile)
            #Else
                pptDoc.ExportAsFixedFormat Path:=tempFile, FixedFormatType:=2, Intent:=2, PrintRange:=Nothing, BitmapMissingFonts:=True
            #End If
    End Select
    
    #If Mac Then
        scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CopyFile", tempFile & APPLE_SCRIPT_SPLIT_KEY & destFile)
    #Else
        If fso.fileExists(tempFile) Then fso.CopyFile tempFile, destFile, True
    #End If
    On Error GoTo 0
    
    #If PRINT_DEBUG_MESSAGES Then
        If Err.Number = 0 Then
            Debug.Print INDENT_LEVEL_2 & "Report saved."
        Else
            Debug.Print INDENT_LEVEL_2 & "Failed to save." & vbNewLine & _
                        INDENT_LEVEL_2 & "Error Number: " & Err.Number & vbNewLine & _
                        INDENT_LEVEL_2 & "Error Description: " & Err.Description
        End If
    #End If
    
    If Val(pptApp.Version) > 15 Then DisableAutoSave pptDoc
    
    WriteFileToDisk = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub DisableAutoSave(ByVal pptDoc As Object)
    On Error Resume Next
    If pptDoc.AutoSaveOn Then pptDoc.AutoSaveOn = False
    On Error GoTo 0
End Sub

Private Sub KillPowerPoint(ByRef pptApp As Object, ByRef pptDoc As Object)
    #If Mac Then
        Dim closeResult As String
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Closing Open PowerPoint Instances"
    #End If
    
    On Error Resume Next
    If Not pptDoc Is Nothing Then
        pptDoc.Close SaveChanges:=False
        Set pptDoc = Nothing
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print INDENT_LEVEL_1 & "Attempting to close the template." & vbNewLine & _
                        INDENT_LEVEL_1 & "Status: " & (pptDoc Is Nothing)
        #End If
    End If
    
    If Not pptApp Is Nothing Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print INDENT_LEVEL_1 & "Attempting to close PowerPoint."
        #End If
        pptApp.Quit
        Set pptApp = Nothing
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print INDENT_LEVEL_1 & "Status: " & (pptApp Is Nothing)
        #End If
    End If

    #If Mac Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print INDENT_LEVEL_1 & "Attempting extra step required to completely close MS PowerPoint on MacOS."
        #End If
    
        closeResult = AppleScriptTask(APPLE_SCRIPT_FILE, "ClosePowerPoint", closeResult)

        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print INDENT_LEVEL_1 & "Status: " & closeResult
        #End If
    #End If
    On Error GoTo 0
End Sub

Private Sub ZipReports(ByVal ws As Worksheet, ByVal savePath As Variant, ByRef saveResult As Boolean, ByVal resourcesFolder As String)
    Dim zipCommand As String
    Dim zipPath As Variant
    Dim zipName As Variant
    Dim zipFileNameElements As Variant
    Dim errDescription As String
    Dim archiverPath As String

    #If Mac Then
        Dim scriptResultString As String
        Dim scriptResultBoolean As Boolean
    #Else
        Dim fso As Object
        Dim shellApp As Object
        Dim archiverName As String
        Dim startTime As Double
    #End If
    
    On Error Resume Next
    If Right$(savePath, 1) <> Application.PathSeparator Then
        savePath = savePath & Application.PathSeparator
    End If
    
    ' Step 1: Build the filename
    zipFileNameElements = ws.Range("C2:C5").Value
    zipName = zipFileNameElements(2, 1) & " (" & _
              zipFileNameElements(1, 1) & " - " & _
              zipFileNameElements(3, 1) & " " & _
              zipFileNameElements(4, 1) & ").zip"
    
    ' Step 2: Set the filepath and remove old copy if present
    #If Mac Then
        zipPath = savePath & zipName
    #Else
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        zipPath = GetTempFilePath(zipName)
        If fso.fileExists(zipPath) Then fso.DeleteFile zipPath, True
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Zipping Reports" & vbNewLine & _
                    INDENT_LEVEL_1 & "Zip Filename: " & zipName & vbNewLine & _
                    INDENT_LEVEL_1 & "Destination:  " & savePath
    #End If
    
    #If Mac Then
        ' Step 3a: Identify archive tool
        archiverPath = FindPathToArchiveTool(resourcesFolder)
            
        ' Step 3b: Generate the zip file
        If archiverPath = "" Then
            scriptResultString = AppleScriptTask(APPLE_SCRIPT_FILE, "CreateZipWithDefaultArchiver", savePath & APPLE_SCRIPT_SPLIT_KEY & zipPath)
        Else
            zipCommand = Chr(34) & archiverPath & Chr(34) & " a " & Chr(34) & zipPath & Chr(34) & " " & Chr(34) & savePath & "*.pdf" & Chr(34)
            scriptResultString = AppleScriptTask(APPLE_SCRIPT_FILE, "CreateZipWithLocal7Zip", zipCommand)
        End If
        
        ' Step 4: Report the result and remove PDFs if successful
        saveResult = (scriptResultString = "Success")
        If saveResult Then
            scriptResultBoolean = AppleScriptTask(APPLE_SCRIPT_FILE, "ClearPDFsAfterZipping", savePath)
        Else
            errDescription = scriptResultString
        End If
    #Else
        ' Step 3a: Identify archive tool
        archiverPath = FindPathToArchiveTool(resourcesFolder, archiverName)
        
        ' Step 3b: Generate the zip file
        Select Case archiverName
            Case "7Zip", "Local 7zip"
                zipCommand = """" & archiverPath & """ a """ & zipPath & """ """ & savePath & "*.pdf" & """"
                Shell zipCommand, vbNormalFocus
            Case Else
                Set shellApp = CreateObject("Shell.Application")
                
                ' Step 3c: Simplify filename in case Hangul support isn't enabled
                zipName = zipFileNameElements(2, 1) & " (" & _
                          zipFileNameElements(3, 1) & " " & _
                          zipFileNameElements(4, 1) & ").zip"
                zipPath = GetTempFilePath(zipName)
        
                ' Step 3d: Check for old copy with the new file name and remove if found
                If fso.fileExists(zipPath) Then fso.DeleteFile zipPath, True
                
                ' Step 4a: Create an empty ZIP file
                Open zipPath For Output As #1
                Print #1, "PK" & Chr$(5) & Chr$(6) & String(18, vbNullChar)
                Close #1
                                
                ' Step 4b: Add the contents of savePath to the zip file
                shellApp.Namespace(zipPath).CopyHere shellApp.Namespace(savePath).Items
        End Select
        
        ' Step 5: Give the system some time to complete the zip process
        startTime = Timer
        Do ' Wait up tp 10 seconds for creation
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop While Not fso.fileExists(zipPath) And Timer - startTime < 10
        
        ' Step 6: Copy the zip file and report if process was successful
        If fso.fileExists(zipPath) Then
            Application.Wait (Now + TimeValue("0:00:02")) ' Wait a couple seconds for the file to be released
            fso.CopyFile zipPath, savePath & zipName, True
            Kill zipPath
            saveResult = True
        Else
            If Err.Number <> 0 Then errDescription = Err.Description
            saveResult = False
        End If
        
        ' Step 7: Remove the PDFs only if the zip was successfully created
        If saveResult Then DeletePDFs savePath
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        If saveResult Then
            Debug.Print INDENT_LEVEL_1 & "Zip successful"
        Else
            Debug.Print INDENT_LEVEL_1 & "Zip failed" & vbNewLine & _
                        INDENT_LEVEL_2 & "Error: " & errDescription
        End If
    #End If
    
    On Error GoTo 0
End Sub
