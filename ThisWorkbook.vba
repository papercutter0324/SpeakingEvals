Option Explicit

Const PRINT_DEBUG_MESSAGES As Boolean = True
Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
Const IS_WORD_VISIBLE As Boolean = True
Const SHOW_WORD_SCREEN_UPDATING As Boolean = True
Dim isAppleScriptInstalled As Boolean

Sub PrintReports()
    Dim ws As Worksheet, wordApp As Object, wordDoc As Object
    Dim templatePath As String, savePath As String, fileName As String
    Dim resultMsg As String, msgToDisplay As String, msgTitle As String, msgType As Integer
    Dim currentRow As Long, lastRow As Long, firstStudentRecord As Integer
    Dim saveResult As Boolean

    Const ERR_INCOMPLETE_RECORDS As String = "incompleteRecords"
    Const ERR_LOADING_WORD As String = "loadingWord"
    Const ERR_LOADING_TEMPLATE As String = "loadingTemplate"
    Const ERR_MISSING_SHAPES As String = "missingTemplateShapes"
    Const MSG_SUCCESS As String = "exportSuccessful"
    Const MSG_FAILED As String = "exportFailed"
    
    #If Mac Then
        RequestInitialFileAndFolderAccess
        isAppleScriptInstalled = CheckForAppleScript()
    #End If
    
    Set ws = ActiveSheet
    If ws Is Nothing Then GoTo Cleanup
    
    If Not VerifyRecordsAreComplete(ws, lastRow, firstStudentRecord) Then
        resultMsg = ERR_INCOMPLETE_RECORDS
        GoTo Cleanup
    End If
    
    templatePath = LoadTemplate()
    If templatePath = "" Then GoTo Cleanup
    
    savePath = SetSaveLocation(ws)
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
        WriteReport ws, wordApp, wordDoc, currentRow, savePath, fileName, saveResult
    Next currentRow
    
    ws.Activate ' Ensure the right worksheet is being shown when finished.
    resultMsg = IIf(saveResult, MSG_SUCCESS, MSG_FAILED)
    
Cleanup:
    Select Case resultMsg
        Case ERR_INCOMPLETE_RECORDS
            msgToDisplay = "One or more fields for missing. Please complete all fields and try again."
            msgTitle = "Missing Data!"
            msgType = vbExclamation
        Case ERR_LOADING_WORD, ERR_LOADING_TEMPLATE
            msgToDisplay = "There was an error opening MS Word and/or the template. This is sometimes normal MS Office behaviour, so please wait a couple seconds and try again."
            msgTitle = "Error!"
            msgType = vbExclamation
        Case ERR_MISSING_SHAPES
            msgToDisplay = "There is a error with the template. Please redownload a copy of the original and try again."
            msgTitle = "Error!"
            msgType = vbExclamation
        Case MSG_SUCCESS
            msgToDisplay = "Export complete!"
            msgTitle = "Process complete!"
            msgType = vbInformation
        Case MSG_FAILED
            msgToDisplay = "Export failed. Please ensure all data was entered correctly and try saving to a different folder."
            msgTitle = "Process failed!"
            msgType = vbInformation
    End Select
    
    If resultMsg <> "" Then MsgBox msgToDisplay, msgType, msgTitle
    KillWord wordApp, wordDoc, ws
End Sub

Private Function LoadWord(ByRef wordApp As Object, ByRef wordDoc As Object, ByVal templatePath As String) As Boolean
    ' Check for an open instance of Word
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    On Error GoTo ErrorHandler
    
    ' Open a new one if not found
    #If Mac Then
        Dim appleScriptResult As String
        appleScriptResult = "N/A"
        
        If isAppleScriptInstalled Then
            appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "LoadApplication", "Microsoft Word")
            If PRINT_DEBUG_MESSAGES Then Debug.Print appleScriptResult
            
            appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "IsAppLoaded", "Microsoft Word")
            If PRINT_DEBUG_MESSAGES Then Debug.Print appleScriptResult
            
            Set wordApp = GetObject(, "Word.Application")
        End If
    #End If
    
    'For Windows users and MacOS users without SpeakingEvals.scpt
    If wordApp Is Nothing Then Set wordApp = CreateObject("Word.Application")
    
    ' Make the process visible so users understand their computer isn't frozen
    wordApp.Visible = IS_WORD_VISIBLE
    wordApp.ScreenUpdating = SHOW_WORD_SCREEN_UPDATING
    
    If Not wordApp Is Nothing Then Set wordDoc = wordApp.Documents.Open(templatePath)
    LoadWord = (Not wordApp Is Nothing)
    Exit Function
ErrorHandler:
    If Err.Number = 0 Then
        Err.Clear
        Resume Next
    Else
        #If Mac Then
            MsgBox "An error occurred while trying to load Microsoft Word. This is usually a result of a quirk in MacOS. Try creating the reports again, and it should work fine." & vbNewLine & vbNewLine & _
            "If the problem persists, please take a picture of the following error message and ask your team leader to send it to Warren at Bundang." & vbNewLine & vbNewLine & _
            "VBA Error " & Err.Number & ": " & Err.Description & vbNewLine & _
            "AppleScript Error: " & appleScriptResult, vbCritical, "Error Loading Word"
        #End If
        
        LoadWord = False
        If Not wordApp Is Nothing Then Set wordApp = Nothing
    End If
End Function

Private Sub KillWord(ByRef wordApp As Object, ByRef wordDoc As Object, ByRef ws As Worksheet)
    If Not wordDoc Is Nothing Then wordDoc.Close SaveChanges:=False
    If Not wordApp Is Nothing Then wordApp.Quit
    
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Set ws = Nothing
    
    #If Mac Then
        Dim closeResult As String
        
        If isAppleScriptInstalled Then
            closeResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CloseWord", closeResult)
            If PRINT_DEBUG_MESSAGES Then Debug.Print closeResult
        End If
    #End If
End Sub

Private Function VerifyRecordsAreComplete(ByRef ws As Worksheet, ByRef lastRow As Long, ByRef firstStudentRecord As Integer) As Boolean
    Dim currentRow As Integer, currentColumn As Integer
    Dim missingData As Boolean
    
    Const CLASS_INFO_FIRST_ROW As Integer = 1
    Const CLASS_INFO_LAST_ROW As Integer = 6
    Const STUDENT_INFO_FIRST_ROW As Integer = 8
    Const STUDENT_INFO_FIRST_COL As Integer = 2
    Const STUDENT_INFO_LAST_COL As Integer = 10
    
    ' Set here are pass back to calling sub in order to keep declarations organized
    firstStudentRecord = STUDENT_INFO_FIRST_ROW
    
    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    On Error GoTo 0
    
    If lastRow < STUDENT_INFO_FIRST_ROW Then
        MsgBox "No students were found!", vbExclamation, "Error!"
        VerifyRecordsAreComplete = False
        Exit Function
    End If
    
    ' Ensure class details are complete
    For currentRow = CLASS_INFO_FIRST_ROW To CLASS_INFO_LAST_ROW
        If PRINT_DEBUG_MESSAGES Then
            Debug.Print ws.Cells(currentRow, 1).Value & " " & ws.Cells(currentRow, 3).Value
        End If
        If IsEmpty(ws.Cells(currentRow, 3).Value) Then
            missingData = True
            Exit For
        End If
    Next currentRow
    
    If missingData Then
        VerifyRecordsAreComplete = False
        Exit Function
    End If
    
    ' Ensure all student records are complete
    For currentRow = STUDENT_INFO_FIRST_ROW To lastRow
        For currentColumn = STUDENT_INFO_FIRST_COL To STUDENT_INFO_LAST_COL
            If PRINT_DEBUG_MESSAGES Then
                Debug.Print ws.Cells(currentRow, 2).Value & "'s " & ws.Cells(7, currentColumn).Value & ": " & ws.Cells(currentRow, currentColumn).Value
            End If
            
            If IsEmpty(ws.Cells(currentRow, currentColumn).Value) Then
                missingData = True
                Exit For
            End If
        Next currentColumn
        If missingData Then Exit For
    Next currentRow

    VerifyRecordsAreComplete = (Not missingData)
End Function

Private Function LoadTemplate() As String
    Dim selectedPath As String, templatePath As String, tempTemplatePath As String, finalTemplatePath As String
    Dim msgToDisplay As String, msgTitle As String
    Dim validTemplateFound As Boolean
    
    Const REPORT_TEMPLATE As String = "Speaking Evaluation Template.docx"
    
    selectedPath = ThisWorkbook.Path
    ConvertOneDriveToLocalPath selectedPath
    
    templatePath = selectedPath & Application.PathSeparator & REPORT_TEMPLATE
    finalTemplatePath = templatePath
    
    #If Mac Then
        tempTemplatePath = Environ("TMPDIR") & REPORT_TEMPLATE
    #Else
        tempTemplatePath = Environ("TEMP") & Application.PathSeparator & REPORT_TEMPLATE
    #End If
    
    DeleteTempTemplate tempTemplatePath
    
    ' Check for the report template and download it if not found
    #If Mac Then
        If isAppleScriptInstalled Then
            validTemplateFound = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", templatePath)
            
            If validTemplateFound Then validTemplateFound = VerifyTemplateHash(templatePath)
            
            ' Work with a temp copy to avoid bugs caused by cloud apps not making the template available quick enough
            If Not validTemplateFound Then
                templatePath = tempTemplatePath
                validTemplateFound = DownloadReportTemplate(templatePath)
            End If
            
            If Not validTemplateFound Then templatePath = ""
        Else
            If Dir(templatePath) = "" Then templatePath = ""
        End If
    #Else
        If Dir(templatePath) <> "" Then
            validTemplateFound = VerifyTemplateHash(templatePath)
        End If
            
        ' Work with a temp copy to avoid bugs caused by cloud apps not making the template available quick enough
        If Not validTemplateFound Then
            templatePath = tempTemplatePath
            validTemplateFound = DownloadReportTemplate(templatePath)
        End If
        
        If Not validTemplateFound Then templatePath = ""
    #End If
        
    If templatePath = "" Then
        msgToDisplay = "No template was found. Process canceled."
        msgTitle = "Template Not Found"
        MsgBox msgToDisplay, vbExclamation, msgTitle
        templatePath = ""
    End If
    
    If PRINT_DEBUG_MESSAGES Then
        Debug.Print IIf(templatePath <> "", "Template loaded: " & templatePath, "No template was found or download failed.")
    End If
    
    If templatePath = tempTemplatePath Then MoveFile tempTemplatePath, finalTemplatePath
    LoadTemplate = templatePath
End Function

Private Sub DeleteTempTemplate(ByVal tempTemplatePath As String)
    #If Mac Then
        Dim appleScriptResult As Boolean
        
        If isAppleScriptInstalled Then
            appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", tempTemplatePath)
            
            If AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", tempTemplatePath) Then
                appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", tempTemplatePath)
            End If
        Else
            Kill tempTemplatePath
        End If
    #Else
        Dim fso As Object
        
        On Error Resume Next
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(tempTemplatePath) Then fso.DeleteFile tempTemplatePath, True
        
        tempTemplatePath = Replace(tempTemplatePath, " ", "%20")
        If fso.FileExists(tempTemplatePath) Then fso.DeleteFile tempTemplatePath, True
        
        Set fso = Nothing
        On Error GoTo 0
    #End If
End Sub

Private Sub MoveFile(ByVal tempTemplatePath As String, ByVal finalTemplatePath As String)
    #If Mac Then
        Dim scriptResult As Boolean, scriptParam As String
        
        If isAppleScriptInstalled Then
            scriptParam = tempTemplatePath & "," & finalTemplatePath
            scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CopyFile", scriptParam)
        Else
            On Error Resume Next
            Name tempTemplatePath As finalTemplatePath
            If PRINT_DEBUG_MESSAGES Then
                If Err.Number <> 0 Then Debug.Print Err.Number & " - " & Err.Description
            End If
            On Error GoTo 0
        End If
    #Else
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFile tempTemplatePath, finalTemplatePath
        Set fso = Nothing
    #End If
End Sub

Private Function DownloadReportTemplate(ByVal filepath As String) As Boolean
    Const REPORT_TEMPLATE_URL As String = "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/Speaking%20Evaluation%20Template.docx"
    Dim downloadCommand As String, downloadResult As Boolean
    
    downloadResult = False
    
    #If Mac Then
        On Error Resume Next
        downloadResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DownloadFile", filepath & "," & REPORT_TEMPLATE_URL)
        
        If downloadResult Then RequestAdditionalFileAndFolderAccess filepath
        
        If PRINT_DEBUG_MESSAGES Then
            Debug.Print IIf(Err.Number = 0, "Download successful.", "Error: " & Err.Description)
        End If
        On Error GoTo 0
    #Else
        Dim objShell As Object, fso As Object, xmlHTTP As Object, fileStream As Object
        Dim useCurl As Boolean, useDotNet35 As Boolean
        
        useCurl = CheckForCurl()
        
        If Not useCurl Then useDotNet35 = CheckForDotNet35()
        
        If useCurl Then
            On Error Resume Next
            Set objShell = CreateObject("WScript.Shell")
            Set fso = CreateObject("Scripting.FileSystemObject")
            
            downloadCommand = "cmd /c curl.exe -o """ & filepath & """ """ & REPORT_TEMPLATE_URL & """"
            objShell.Run downloadCommand, 0, True
            downloadResult = fso.FileExists(filepath)
            
            Set objShell = Nothing
            Set fso = Nothing
            On Error GoTo 0
        ElseIf useDotNet35 Then
            Set xmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
            Set fileStream = CreateObject("ADODB.Stream")
            
            xmlHTTP.Open "Get", REPORT_TEMPLATE_URL, False
            xmlHTTP.Send
            
            If xmlHTTP.Status = 200 Then
                fileStream.Open
                fileStream.Type = 1 ' Binary
                fileStream.Write xmlHTTP.responseBody
                fileStream.SaveToFile filepath, 2 ' Overwrite existing, if somehow present
                fileStream.Close
                downloadResult = True
            End If
            
            Set xmlHTTP = Nothing
            Set fileStream = Nothing
        End If
    #End If
    
    If downloadResult Then
        DownloadReportTemplate = VerifyTemplateHash(filepath)
    Else
        DownloadReportTemplate = False
    End If
End Function

Private Function VerifyTemplateHash(ByVal filepath As String) As Boolean
    Dim md5Command As String, generatedHash As String
    
    Const TEMPLATE_HASH As String = "1D40D1790DCE2C5AA405A05BDA981517"
    
    #If Mac Then
        If Not isAppleScriptInstalled Then
            MsgBox "SpeakingEvals.scpt has not been installed, so the report template's file integrity cannot be validated." & vbNewLine & vbNewLine & _
                   "The reports will still be created, but please check that everything looks okay or download the template manually." & vbNewLine & _
                   vbNewLine & Space(40) & "Press Ok to continue."
            VerifyTemplateHash = True
            Exit Function
        End If
            
        On Error Resume Next
        VerifyTemplateHash = AppleScriptTask(APPLE_SCRIPT_FILE, "CompareMD5Hashes", filepath & "," & TEMPLATE_HASH)
        If Err.Number = 0 Then
            Exit Function
        Else
            GoTo ErrorHandler
        End If
    #Else
        Dim objShell As Object, fso As Object
        Dim tempHashFile As String
    
        On Error GoTo ErrorHandler
        tempHashFile = Environ("TEMP") & "\hash_result.txt"
        md5Command = "cmd /c certutil -hashfile """ & filepath & """ MD5 > """ & tempHashFile & """"
        
        ' Get the hash of the downloaded template
        Set objShell = CreateObject("WScript.Shell")
        objShell.Run md5Command, 0, True
        
        ' Load the hash into memory and delete hash_result.txt
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(tempHashFile) Then
            With fso.OpenTextFile(tempHashFile, 1)
                .ReadLine
                generatedHash = Trim(LCase(.ReadLine))
                .Close
            End With
            fso.DeleteFile tempHashFile
        Else
            VerifyTemplateHash = False
            Exit Function
        End If
        On Error GoTo 0
        
        Set objShell = Nothing
        Set fso = Nothing
        
        VerifyTemplateHash = (LCase(TEMPLATE_HASH) = generatedHash)
        Exit Function
    #End If
ErrorHandler:
    If PRINT_DEBUG_MESSAGES Then Debug.Print "Error: " & Err.Number & " - " & Err.Description
    
    VerifyTemplateHash = False
    
    #If Mac Then
    #Else
        If Not objShell Is Nothing Then Set objShell = Nothing
        If Not fso Is Nothing Then Set fso = Nothing
    #End If
End Function

Private Function SetSaveLocation(ByRef ws As Object) As String
    Dim selectedPath As String
    Dim scriptResult As Boolean
    
    selectedPath = ThisWorkbook.Path & Application.PathSeparator & GenerateSaveFolderName(ws) & Application.PathSeparator
    ConvertOneDriveToLocalPath selectedPath
    
    If PRINT_DEBUG_MESSAGES Then Debug.Print "Saving reports in: " & vbNewLine & "    " & selectedPath; ""
    
    #If Mac Then
        Dim folderAction As String
        
        If isAppleScriptInstalled Then
            scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFolderExist", selectedPath)
            
            ' If the folder exists, delete it to avoid confusion
            If scriptResult Then
                ' This MUST occur before falling back to using the same directory as the Excel file.
                scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "ClearFolder", selectedPath)
            End If
            
            scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CreateFolder", selectedPath)
            RequestAdditionalFileAndFolderAccess selectedPath
        End If
    #End If
    
    ' Handle Windows users and MacOS users who chose not to installed SpeakingEvals.scpt
    If Not isAppleScriptInstalled And Not scriptResult Then
        If Dir(selectedPath, vbDirectory) <> "" Then
            If PRINT_DEBUG_MESSAGES Then Debug.Print "Directory already exists. Clearing contents."
            DeleteExistingFolder selectedPath
        Else
            If PRINT_DEBUG_MESSAGES Then Debug.Print "Path not found. Attempting to create."
        End If
        
        CreateSaveFolder selectedPath
    End If
    
    ConvertOneDriveToLocalPath selectedPath 'This appears no longer necessary
    SetSaveLocation = selectedPath
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
        Case "TTh"
            classIdentifier = "TTh - " & ws.Cells(5, 3).Value
        Case "MWF (Class 1)": classIdentifier = "MWF-1"
        Case "MWF (Class 2)": classIdentifier = "MWF-2"
        Case "TTh (Class 1)": classIdentifier = "TTh-1"
        Case "TTh (Class 2)": classIdentifier = "TTh-2"
    End Select
    
    GenerateSaveFolderName = ws.Cells(3, 3).Value & " (" & classIdentifier & ")"
End Function

Private Sub CreateSaveFolder(ByRef selectedPath As String)
    If Right(selectedPath, 1) = Application.PathSeparator Then
        selectedPath = Left(selectedPath, Len(selectedPath) - 1)
    End If
    
    #If Mac Then
        Dim msgToDisplay As String, msgTitle As String
        
        On Error Resume Next
        If Dir(selectedPath, vbDirectory) = "" Then
            MkDir selectedPath
            
            If Err.Number <> 0 Then
                If PRINT_DEBUG_MESSAGES Then Debug.Print "Error creating directory: " & Err.Description
            
                msgToDisplay = "Unable to create separate report folders. Student reports will be saved in the same location as this Excel file."
                msgTitle = "Notice"
                MsgBox msgToDisplay, vbExclamation, msgTitle
                selectedPath = ThisWorkbook.Path & Application.PathSeparator
            Else
                If PRINT_DEBUG_MESSAGES Then Debug.Print "Directories successfully created. Continuing process."
            End If
        End If
          
        If Not isAppleScriptInstalled Then
            ' Verify the user has deleted any existing files. If files are discovered, warn the user that any files with
            ' matching filenames will be overwritten and existing files may be mixed in with the newly created ones.
            If Dir(selectedPath & "/*") <> "" Then
                msgToDisplay = "It appears some files still exist in """ & selectedPath & """. " & vbNewLine & vbNewLine & "The new reports will be generated, but " & _
                               "any existing files with the same filenames will be overwritten, and any existing files will be mixed in with the newly generated ones."
                msgTitle = "Notice"
                MsgBox msgToDisplay, vbExclamation, msgTitle
            End If
        End If
                
        RequestAdditionalFileAndFolderAccess selectedPath
    #Else
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder selectedPath
        Set fso = Nothing
    #End If
    
    If Right(selectedPath, 1) <> Application.PathSeparator Then
        selectedPath = selectedPath & Application.PathSeparator
    End If
End Sub

Private Sub DeleteExistingFolder(ByVal selectedPath As String)
    #If Mac Then
        Dim msgToDisplay As String, msgTitle As String
        
        msgToDisplay = "Because " & APPLE_SCRIPT_FILE & " is not installed, Excel is unable to delete any existing reports for this class. It is recommended to delete them before continuing." & _
                       vbNewLine & vbNewLine & "You can safely delete any .pdf files in """ & selectedPath & """ now and then click 'Okay' to continue."
        msgTitle = "Notice"
        MsgBox msgToDisplay, vbExclamation, msgTitle
    #Else
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        If Right(selectedPath, 1) = Application.PathSeparator Then
            selectedPath = Left(selectedPath, Len(selectedPath) - 1)
        End If
        
        fso.DeleteFolder selectedPath, True
        Set fso = Nothing
    #End If
End Sub

Private Sub WriteReport(ByRef ws As Object, ByRef wordApp As Object, ByRef wordDoc As Object, ByVal currentRow As Integer, ByVal savePath As String, ByRef fileName As String, ByRef saveResult As Boolean)
    Dim nativeTeacher As String, koreanTeacher As String, classLevel As String, classTime As String, evalDate As String
    Dim englishName As String, koreanName As String, grammarScore As String, pronunciationScore As String, fluencyScore As String
    Dim mannerScore As String, contentScore As String, effortScore As String, commentText As String, overallGrade As String
    Dim signatureAdded As Boolean
    
    ' Data applicable to all reports
    nativeTeacher = ws.Cells(1, 3).Value
    koreanTeacher = ws.Cells(2, 3).Value
    classLevel = ws.Cells(3, 3).Value
    classTime = ws.Cells(4, 3).Value & "-" & ws.Cells(5, 3).Value
    evalDate = ws.Cells(6, 3).Value
    evalDate = Format(Date, "MMM. YYYY")
    
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
    
    If PRINT_DEBUG_MESSAGES Then Debug.Print "Saving to: " & savePath & vbNewLine & "Saving as: " & fileName
    
    With wordDoc
        ' Populate the report's header
        .Shapes("Report_Header").GroupItems("English_Name").TextFrame.TextRange.Text = englishName
        .Shapes("Report_Header").GroupItems("Korean_Name").TextFrame.TextRange.Text = koreanName
        .Shapes("Report_Header").GroupItems("Grade").TextFrame.TextRange.Text = classLevel
        .Shapes("Report_Header").GroupItems("Native_Teacher").TextFrame.TextRange.Text = nativeTeacher
        .Shapes("Report_Header").GroupItems("Korean_Teacher").TextFrame.TextRange.Text = koreanTeacher
        .Shapes("Report_Header").GroupItems("Date").TextFrame.TextRange.Text = evalDate
        
        ' Populate the scores
        .Shapes("Grammar_Scores").GroupItems("Grammar_" & grammarScore).TextFrame.TextRange.Text = grammarScore
        .Shapes("Pronunciation_Scores").GroupItems("Pronunciation_" & pronunciationScore).TextFrame.TextRange.Text = pronunciationScore
        .Shapes("Fluency_Scores").GroupItems("Fluency_" & fluencyScore).TextFrame.TextRange.Text = fluencyScore
        .Shapes("Manner_Scores").GroupItems("Manner_" & mannerScore).TextFrame.TextRange.Text = mannerScore
        .Shapes("Content_Scores").GroupItems("Content_" & contentScore).TextFrame.TextRange.Text = contentScore
        .Shapes("Effort_Scores").GroupItems("Effort_" & effortScore).TextFrame.TextRange.Text = effortScore
        
        ' Populate the comment and overall grade
        .Shapes("Report_Footer").GroupItems("Comments").TextFrame.TextRange.Text = commentText
        .Shapes("Report_Footer").GroupItems("Overall_Grade").TextFrame.TextRange.Text = overallGrade
    End With
    
    On Error Resume Next
    ' Quick check to make sure the teacher's signature is only added once
    signatureAdded = (Not wordDoc.Shapes("Signature") Is Nothing)
    On Error GoTo 0
    
    If Not signatureAdded Then InsertSignature wordDoc
    
    On Error Resume Next
    #If Mac Then
        ' Export to PDF is a bit flaky on MacOS, so we need to do a full SaveAs2. Only results in a minimal time loss.
        wordDoc.SaveAs2 fileName:=(savePath & fileName & ".pdf"), FileFormat:=17, AddtoRecentFiles:=False, EmbedTrueTypeFonts:=True
    #Else
        wordDoc.ExportAsFixedFormat OutputFileName:=(savePath & fileName & ".pdf"), ExportFormat:=17, BitmapMissingFonts:=True
    #End If
    
    saveResult = (Err.Number = 0)
    
    If PRINT_DEBUG_MESSAGES Then
        If saveResult Then
            Debug.Print "Successfully saved: " & fileName
        Else
            Debug.Print "Failed to save." & "Error Number: " & Err.Number & vbNewLine & "Error Description: " & Err.Description
        End If
    End If
    On Error GoTo 0
End Sub

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
    
    Select Case avgScore
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
    
    ' Only add if not already present.
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
        .Height = imageHeight
        
        ' Ensure positioning relative to the page
        .RelativeHorizontalPosition = 1
        .RelativeVerticalPosition = 1
    End With
End Sub

Private Sub SaveSignature(ByVal SIGNATURE_SHAPE_NAME As String, ByRef savePath As String)
    Dim signSheet As Worksheet, tempSheet As Worksheet
    Dim signatureshp As Shape, chrt As ChartObject
    
    Set signatureshp = ThisWorkbook.Worksheets("Instructions").Shapes(SIGNATURE_SHAPE_NAME)
    
    #If Mac Then
        savePath = Environ("TMPDIR") & "tempSignature.png"
    #Else
        savePath = Environ("TEMP") & Application.PathSeparator & "tempSignature.png"
    #End If
    
    ConvertOneDriveToLocalPath savePath
    
    Sheets.Add(, Sheets(Sheets.Count)).Name = "Temp_signature"
    Set tempSheet = Sheets("Temp_signature")
    tempSheet.Select
    
    signatureshp.Copy
    
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

Private Function VerifyAllDocShapesExist(ByRef wordDoc As Object) As Boolean
    Dim shp As Shape, shapeNames As Variant
    Dim msgToDisplay As String, i As Integer
    
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
            If PRINT_DEBUG_MESSAGES Then Debug.Print "Missing shape: " & shapeNames(i)
            msgToDisplay = "There is a critical error with the template. Please redownload a copy of the original and try again."
            MsgBox msgToDisplay, vbExclamation, "Error!"
            VerifyAllDocShapesExist = False
            Exit Function
        End If
    Next i
                       
    VerifyAllDocShapesExist = True
End Function

Private Function WordDocShapeExists(ByRef wordDoc As Object, ByVal shapeName As String) As Boolean
    Dim shp As Object, grpItem As Object
    
    If PRINT_DEBUG_MESSAGES Then Debug.Print "Search for shape: " & shapeName
    
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
    
    If PRINT_DEBUG_MESSAGES Then Debug.Print "Unable to find shape: " & shapeName
    
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

#If Mac Then
Private Function CheckForAppleScript() As Boolean
    Dim appleScriptPath As String
    
    appleScriptPath = "/Users/" & Environ("USER") & "/Library/Application Scripts/com.microsoft.Excel/" & APPLE_SCRIPT_FILE
    
    If PRINT_DEBUG_MESSAGES Then Debug.Print "Locating " & APPLE_SCRIPT_FILE & vbNewLine & "Searching: " & appleScriptPath
    
    On Error Resume Next
    CheckForAppleScript = (Dir(appleScriptPath, vbDirectory) = APPLE_SCRIPT_FILE)
    On Error GoTo 0
    
    If PRINT_DEBUG_MESSAGES Then Debug.Print "Found: " & CheckForAppleScript
End Function

Private Sub RequestInitialFileAndFolderAccess()
    Dim workingFolder As String, tempFolder As String
    Dim filePermissionCandidates As Variant
    Dim fileAccessGranted As Boolean
    
    workingFolder = ThisWorkbook.Path
    tempFolder = Environ("TMPDIR")
    
    ConvertOneDriveToLocalPath workingFolder
    ConvertOneDriveToLocalPath tempFolder
    
    If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Requesting access to: " & vbNewLine & _
                    "    " & workingFolder & vbNewLine & _
                    "    " & tempFolder
    End If
    
    filePermissionCandidates = Array(workingFolder, tempFolder)
    fileAccessGranted = GrantAccessToMultipleFiles(filePermissionCandidates)
    
    If PRINT_DEBUG_MESSAGES Then Debug.Print "Access granted: " & fileAccessGranted
End Sub

Private Sub RequestAdditionalFileAndFolderAccess(ByVal newPath As String)
    Dim filePermissionCandidates As Variant
    Dim fileAccessGranted As Boolean
     
    ConvertOneDriveToLocalPath newPath
    filePermissionCandidates = Array(newPath)
    fileAccessGranted = GrantAccessToMultipleFiles(filePermissionCandidates)
End Sub
#Else
Private Function CheckForCurl() As Boolean
    Dim objShell As Object, objExec As Object
    Dim checkResult As Boolean
    Dim output As String
    
    On Error GoTo ErrorHandler
    If PRINT_DEBUG_MESSAGES Then Debug.Print "Checking if curl.exe is installed."
    
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.exec("cmd /c curl.exe --version")
    
    If Not objExec Is Nothing Then
        Do While Not objExec.StdOut.AtEndOfStream
            output = output & objExec.StdOut.ReadLine() & vbNewLine
        Loop
        checkResult = ((InStr(output, "curl")) > 0)
    End If
    
    If PRINT_DEBUG_MESSAGES Then Debug.Print IIf(checkResult, "   Installed.", "   Not installed. Falling back to .Net.")
    
    CheckForCurl = checkResult
Cleanup:
    If Not objExec Is Nothing Then Set objExec = Nothing
    If Not objShell Is Nothing Then Set objShell = Nothing
    Exit Function
ErrorHandler:
    If PRINT_DEBUG_MESSAGES Then Debug.Print "Error while checking for curl.exe: " & Err.Description
    CheckForCurl = False
    Resume Cleanup
End Function

Private Function CheckForDotNet35() As Boolean
    Dim frameworkPath As String
    
    On Error GoTo ErrorHandler
    If PRINT_DEBUG_MESSAGES Then Debug.Print "Verifying that Microsoft DotNet 3.5 is installed."
    
    frameworkPath = Environ$("systemroot") & "\Microsoft.NET\Framework\v3.5"
    CheckForDotNet35 = Dir$(frameworkPath, vbDirectory) <> vbNullString
    
    If PRINT_DEBUG_MESSAGES Then
        Debug.Print "   Checking path: " & frameworkPath & vbNewLine & _
                    "   Installed: " & CheckForDotNet35
    End If
    
    Exit Function
ErrorHandler:
    If PRINT_DEBUG_MESSAGES Then Debug.Print "Error while checking for .NET 3.5: " & Err.Description
    CheckForDotNet35 = False
End Function
#End If
