Option Explicit

#Const PRINT_DEBUG_MESSAGES = True
#If Mac Then
    Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
    Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
#End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Report Generation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReports(ByVal ws As Worksheet, ByVal clickedButtonName As String)
    ' Objects to open PowerPoint and modify the template
    Dim pptApp As Object
    Dim pptDoc As Object
    
    ' Variables for determining the code path and important states
    Dim generateProcess As String
    Dim saveResult As Boolean
    
    ' Strings for generating important messages for the user
    Dim resultMsg As String
    Dim msgToDisplay As String
    Dim msgTitle As String
    Dim msgType As Long
    Dim dialogSize As Long
    Dim msgresult As Variant
    
    ' Strings for tracking important filenames and filepaths
    Dim resourcesFolder As String
    Dim templatePath As String
    Dim savePath As String
    
    ' Numbers for iterating through student records and generate the reports
    Dim currentRow As Long
    Dim lastRow As Long
    Dim firstStudentRecord As Long
    Dim i As Long
    
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
    
    #If Mac Then
        Dim scriptResult As Boolean
    #End If
    
    Select Case clickedButtonName
        Case "Button_GenerateReports"
            generateProcess = "FinalReports"
        Case "Button_GenerateProofs"
            generateProcess = "Proofs"
        Case Else
            msgToDisplay = "You have clicked an invalid option for creating the reports. This shouldn't be possible unless this file has been altered " & _
                           "in an unintended manner. Please download a new copy of this Excel file, copy over all of the students' records, and try again."
            msgresult = DisplayMessage(msgToDisplay, vbExclamation, "Invalid Selection!")
        Exit Sub
    End Select
    
    ' Verify character limits are not exceeded
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Beginning Report Generation" & vbNewLine & _
                    "    Report Type: " & generateProcess
    #End If

    resourcesFolder = ThisWorkbook.Path & Application.PathSeparator & "Resources"
    ConvertOneDriveToLocalPath resourcesFolder
    
    #If Mac Then
        If Not RequestFileAndFolderAccess(resourcesFolder) Then
            ' Create an error msg
            ' GoTo CleanUp
        End If
    #Else
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Checking for Resources Folder" & vbNewLine & _
                        "    Path: " & resourcesFolder
        #End If
        
        If Not DoesFolderExist(resourcesFolder) Then
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "    Folder not found. Attempting to create."
            #End If
            MkDir resourcesFolder
        End If
        
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Folder Present: " & DoesFolderExist(resourcesFolder)
        #End If
        
        If Not DoesFolderExist(resourcesFolder) Then
            resultMsg = ERR_RESOURCES_FOLDER
            GoTo CleanUp
        End If
    #End If
    
    If Not InstallFonts() Then
        resultMsg = ERR_FONT_INSTALLATION
        GoTo CleanUp
    End If

    If IsPptTemplateAlreadyOpen(resourcesFolder, REPORT_TEMPLATE) Then
        ' I can probably set an error msg and send this to CleanUp
        Exit Sub
    End If

    If Not VerifyRecordsAreComplete(ws, lastRow, firstStudentRecord) Then
        resultMsg = ERR_INCOMPLETE_RECORDS
        GoTo CleanUp
    End If

    templatePath = LocateTemplate(resourcesFolder, REPORT_TEMPLATE)
    If templatePath = vbNullString Then
        ' Set an error msg
        GoTo CleanUp
    End If

    savePath = SetSaveLocation(ws, generateProcess, resourcesFolder)
    If savePath = vbNullString Then
        ' Set an error msg
        GoTo CleanUp
    End If

    If Not LoadPowerPoint(pptApp, pptDoc, templatePath) Then
        resultMsg = ERR_LOADING_POWERPOINT
        GoTo CleanUp
    End If
    
    If pptDoc Is Nothing Then
        resultMsg = ERR_LOADING_TEMPLATE
        GoTo CleanUp
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Generate Reports"
    #End If
    
    For currentRow = firstStudentRecord To lastRow
        #If PRINT_DEBUG_MESSAGES Then
            i = i + 1
            Debug.Print "    Generating Report " & i & " of " & (lastRow - firstStudentRecord + 1)
        #End If
        WritePptReport ws, pptApp, pptDoc, generateProcess, currentRow, savePath, saveResult
    Next currentRow
    
    If Not saveResult Then
        resultMsg = MSG_SAVE_FAILED
        GoTo CleanUp
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Save process complete."
    #End If
    
    KillPowerPoint pptApp, pptDoc
    resultMsg = MSG_SUCCESS
    
    If generateProcess = "FinalReports" Then
        ZipReports ws, savePath, saveResult, resourcesFolder
        If Not saveResult Then resultMsg = MSG_ZIP_FAILED
    End If
    
    If saveResult Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Attempting to open destination folder." & vbNewLine & _
                        "    Path: " & savePath
        #End If
        
        #If Mac Then
            msgToDisplay = "Generated reports have been saved to: " & vbNewLine & savePath
            msgTitle = "Notice!"
            msgType = vbInformation
            dialogSize = 350
            msgresult = DisplayMessage(msgToDisplay, msgType, msgTitle, dialogSize)
        #Else
            Shell "explorer.exe """ & savePath & """", vbNormalFocus
        #End If
    End If
    
CleanUp:
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
    
    If resultMsg <> vbNullString Then msgresult = DisplayMessage(msgToDisplay, msgType, msgTitle, dialogSize)
    If Not pptApp Is Nothing Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Beginning final cleanup checks."
        #End If
        KillPowerPoint pptApp, pptDoc
    End If
End Sub

Private Function IsPptTemplateAlreadyOpen(ByVal resourcesFolder As String, ByVal REPORT_TEMPLATE As String) As Boolean
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
            Debug.Print "    PowerPoint instance found" & vbNewLine & _
                        "    Checking if template is already open "
        #End If
        
        templatePath = resourcesFolder & Application.PathSeparator & REPORT_TEMPLATE
        
        For Each pptDoc In pptApp.Presentations
            pathOfOpenDoc = pptDoc.FullName
            ConvertOneDriveToLocalPath pathOfOpenDoc
            If StrComp(pathOfOpenDoc, templatePath, vbTextCompare) = 0 Then
                templateIsOpen = True
                
                #If PRINT_DEBUG_MESSAGES Then
                    Debug.Print "    Open instance found" & vbNewLine & _
                                "    Asking if user wishes to automatically close and continue."
                #End If
                
                msgToDisplay = "An open instance of MS PowerPoint has been detected. Please save any open files before continuing." & vbNewLine & vbNewLine & _
                               "Click OK to automatically close PowerPoint and continue, or click Cancel to finish and save your work."
                               
                If DisplayMessage(msgToDisplay, vbOKCancel + vbCritical, "Notice!", 310) = vbOK Then
                    pptDoc.Close SaveChanges:=False
                    templateIsOpen = False
                    #If PRINT_DEBUG_MESSAGES Then
                        Debug.Print "    Open instance has been closed."
                    #End If
                End If
            End If
        Next pptDoc
    End If
    On Error GoTo 0
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Open instances: " & templateIsOpen
    #End If
    
    Set pptDoc = Nothing
    Set pptApp = Nothing
    IsPptTemplateAlreadyOpen = templateIsOpen
End Function

Private Function LoadPowerPoint(ByRef pptApp As Object, ByRef pptDoc As Object, ByVal templatePath As String) As Boolean
    #If Mac Then
        Dim appleScriptResult As String
        Dim msgToDisplay As String
        Dim msgresult As Variant
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
                Debug.Print "    " & appleScriptResult
            #End If
            
            Set pptApp = GetObject(, "PowerPoint.Application")
        End If
    #Else
        If pptApp Is Nothing Then Set pptApp = CreateObject("PowerPoint.Application")
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    PowerPoint Loaded: " & (Not pptApp Is Nothing)
    #End If
    
    ' Make the process visible so users understand their computer isn't frozen
    pptApp.Visible = True
    
    If Not pptApp Is Nothing Then
        Set pptDoc = pptApp.Presentations.Open(templatePath)
        If val(pptApp.Version) > 15 Then
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "    Disabling AutoSave"
            #End If
            DisableAutoSave pptDoc
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "    Status: " & pptDoc.AutoSaveOn
            #End If
        End If
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Template loaded: " & (Not pptDoc Is Nothing)
    #End If
    
    LoadPowerPoint = (Not pptApp Is Nothing)
    Exit Function
ErrorHandler:
    #If Mac Then
        msgToDisplay = "An error occurred while trying to load Microsoft PowerPoint. This is usually a result of a quirk in MacOS. Try creating the reports again, and it should work fine." & vbNewLine & vbNewLine & _
                        "If the problem persists, please take a picture of the following error message and ask your team leader to send it to Warren at Bundang." & vbNewLine & vbNewLine & _
                        "VBA Error " & Err.Number & ": " & Err.Description & vbNewLine & "AppleScript Error: " & appleScriptResult
        msgresult = DisplayMessage(msgToDisplay, vbOKOnly, "Error Loading PowerPoint", 470)
    #End If
    LoadPowerPoint = False
End Function

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
        Debug.Print "        Preparing Report Data"
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
    evalDate = Format$(CDate(reportMetaData(6, 1)), "MMM. YYYY")
    
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
        Debug.Print "        Filename: " & fileName
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
        If .Item("Signature") Is Nothing Then InsertSignature pptDoc
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
        If Mid(englishName, i, 1) = " " Then
            posToReplace = i
            Exit For
        End If
    Next i
    
    ' Step 3: Update englishName
    If posToReplace <> -1 Then
        ' Step 3a: Replace a space with vbCrLf and split into two lines
        englishName = Left(englishName, posToReplace - 1) & vbCrLf & Mid(englishName, posToReplace + 1)
    Else
        ' Step 3b: Hyphenate and split into two lines
        englishName = Left(englishName, lowerBound) & "-" & vbCrLf & Mid(englishName, lowerBound + 1)
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
    signaturePath = ThisWorkbook.Path & Application.PathSeparator
    ConvertOneDriveToLocalPath signaturePath
    useEmbeddedSignature = (Not mySignature.Shapes.[_Default](SIGNATURE_SHAPE_NAME) Is Nothing)
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
    
    Set signatureshp = mySignature.Shapes.[_Default](SIGNATURE_SHAPE_NAME)
    signatureshp.Copy
    
    signatureImagePath = GetTempFilePath("tempSignature.png")
    ConvertOneDriveToLocalPath signatureImagePath
    
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
        Case Else
            #If Mac Then
                scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "SavePptAsPdf", tempFile)
            #Else
                pptDoc.ExportAsFixedFormat Path:=tempFile, FixedFormatType:=2, Intent:=1, PrintRange:=Nothing, BitmapMissingFonts:=True
            #End If
    End Select
    
    #If Mac Then
        scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CopyFile", tempFile & APPLE_SCRIPT_SPLIT_KEY & destFile)
    #Else
        If fso.FileExists(tempFile) Then fso.CopyFile tempFile, destFile, True
    #End If
    On Error GoTo 0
    
    #If PRINT_DEBUG_MESSAGES Then
        If Err.Number = 0 Then
            Debug.Print "        Report saved."
        Else
            Debug.Print "        Failed to save." & vbNewLine & _
                        "        Error Number: " & Err.Number & vbNewLine & _
                        "        Error Description: " & Err.Description
        End If
    #End If
    
    If val(pptApp.Version) > 15 Then DisableAutoSave pptDoc
    
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
            Debug.Print "    Attempting to close the template." & vbNewLine & _
                        "    Status: " & (pptDoc Is Nothing)
        #End If
    End If
    
    If Not pptApp Is Nothing Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Attempting to close PowerPoint."
        #End If
        pptApp.Quit
        Set pptApp = Nothing
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Status: " & (pptApp Is Nothing)
        #End If
    End If

    #If Mac Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Attempting extra step required to completely close MS PowerPoint on MacOS."
        #End If
    
        closeResult = AppleScriptTask(APPLE_SCRIPT_FILE, "ClosePowerPoint", closeResult)

        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Status: " & closeResult
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
        If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Zipping Reports" & vbNewLine & _
                    "    Zip Filename: " & zipName & vbNewLine & _
                    "    Destination:  " & savePath
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
                If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True
                
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
        Loop While Not fso.FileExists(zipPath) And Timer - startTime < 10
        
        ' Step 6: Copy the zip file and report if process was successful
        If fso.FileExists(zipPath) Then
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
            Debug.Print "    Zip successful"
        Else
            Debug.Print "    Zip failed" & vbNewLine & _
                        "        Error: " & errDescription
        End If
    #End If
    
    On Error GoTo 0
End Sub
