Option Explicit

Const printDebugMessages As Boolean = False
Const isWordAppVisible As Boolean = True
Const showWordAppScreenUpdating As Boolean = True
Dim isAppleScriptInstalled As Boolean

Sub PrintReports()
    Dim ws As Worksheet, wordApp As Object, wordDoc As Object
    Dim templatePath As String, savePath As String, fileName As String
    Dim msgToDisplay As String, msgTitle As String, dbgMsg As String
    Dim currentRow As Long, lastRow As Long
    Dim saveResult As Boolean
    
    ' Disable until code can be reviewed and tested
    isAppleScriptInstalled = False
    '#If Mac Then
    '    isAppleScriptInstalled = checkForAppleScript()
    '    If Not isAppleScriptInstalled Then
    '        PromptToInstallAppleScript
    '    End If
    '#Else
    '    isAppleScriptInstalled = True
    '#End If
    
    ' Initialize Word early so it has time to load.
    ' Doesn't fully solve the problem, but it's an improvement
    If Not LoadWord(wordApp) Then Exit Sub
    
    Set ws = ActiveSheet

    If ws Is Nothing Then
        If printDebugMessages Then
            Debug.Print "Error selecting worksheet!"
        End If
        
        KillWord wordApp, wordDoc, ws
        Exit Sub
    End If
    
    If Not VerifyRecordsAreComplete(ws, lastRow) Then
        KillWord wordApp, wordDoc, ws
        Exit Sub
    End If
    
    templatePath = LoadTemplate()
    If templatePath = "" Then
        KillWord wordApp, wordDoc, ws
        Exit Sub
    End If
    
    savePath = SetSaveLocation(ws)
    If savePath = "" Then
        KillWord wordApp, wordDoc, ws
        Exit Sub
    End If
    
    #If Mac Then
        RequestMacExcelPermissions templatePath, savePath
    #End If
    
    Set wordDoc = wordApp.Documents.Open(templatePath)
    
    If wordDoc Is Nothing Then
        msgToDisplay = "There was an error loading the template. Please wait a couple seconds and try again."
        msgTitle = "Error!"
        MsgBox msgToDisplay, vbExclamation, msgTitle
        KillWord wordApp, wordDoc, ws
        Exit Sub
    End If
    
    If Not VerifyAllDocShapesExist(wordDoc) Then
        msgToDisplay = "There is a critical error with the template. Please redownload a copy of the original and try again."
        msgTitle = "Error!"
        MsgBox msgToDisplay, vbExclamation, msgTitle
        KillWord wordApp, wordDoc, ws
        Exit Sub
    End If
    
    For currentRow = 8 To lastRow
        ClearAllTextBoxes wordDoc
        WriteReport ws, wordApp, wordDoc, currentRow, savePath, fileName, saveResult
    Next currentRow
    
    KillWord wordApp, wordDoc, ws
    
    If saveResult Then
        msgToDisplay = "Export complete!"
        msgTitle = "Process complete!"
    Else
        msgToDisplay = "Export failed. Please ensure all data was entered correctly and try saving to a different folder."
        msgTitle = "Process failed!"
    End If
    
    MsgBox msgToDisplay, vbInformation, msgTitle
End Sub

Private Function LoadWord(ByRef wordApp As Object) As Boolean
    On Error Resume Next
    ' Get a reference to an already running instance of Word
    Set wordApp = GetObject(, "Word.Application")
    
    ' or open a new one if not found
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    
    ' Add a short wait to give Word to load on MacOS
    #If Mac Then
        WaitTimer 3
    #End If
    
    ' Make sure these are enabled so users understand their computer isn't frozen
    wordApp.Visible = isWordAppVisible
    wordApp.ScreenUpdating = showWordAppScreenUpdating
    
    LoadWord = (Not wordApp Is Nothing)
End Function

Private Sub KillWord(ByRef wordApp As Object, ByRef wordDoc As Object, ByRef ws As Worksheet)
    If Not wordDoc Is Nothing Then wordDoc.Close SaveChanges:=False
    If Not wordApp Is Nothing Then wordApp.Quit
    
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Set ws = Nothing
    
    #If Mac Then
        If isAppleScriptInstalled Then
            Dim closeResult As String
            closeResult = AppleScriptTask("SpeakingEvals.scpt", "CloseWord", closeResult)
            If printDebugMessages Then
                Debug.Print closeResult
            End If
        End If
    #End If
End Sub

Private Function VerifyRecordsAreComplete(ByRef ws As Worksheet, ByRef lastRow As Long) As Boolean
    Dim msgToDisplay As String, msgTitle As String
    Dim currentRow As Integer, currentColumn As Integer
    Dim missingData As Boolean, missingStudents As String
    Dim studentName As String, columnName As String
    
    missingData = False
    missingStudents = ""
    
    ' Find the last row containing a student's English name
    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    On Error GoTo 0
    
    ' Student records start in row 8
    If lastRow < 8 Then
        msgToDisplay = "No students were found!"
        msgTitle = "Error!"
        MsgBox msgToDisplay, vbExclamation, msgTitle
        
        VerifyRecordsAreComplete = False
        Exit Function
    End If
    
    ' Verify all fields have been completed
    For currentRow = 8 To lastRow
        For currentColumn = 1 To 9
            If printDebugMessages Then
                studentName = ws.Cells(currentRow, 1).Value
                columnName = ws.Cells(7, currentColumn).Value
                Debug.Print studentName & "'s " & columnName & ": " & ws.Cells(currentRow, currentColumn).Value
            End If
            
            If IsEmpty(ws.Cells(currentRow, currentColumn).Value) Then
                missingData = True
                msgToDisplay = "One or more fields for missing. Please complete all fields and try again."
                msgTitle = "Missing Data!"
                MsgBox msgToDisplay, vbExclamation, msgTitle
                Exit For
            End If
        Next currentColumn
        If missingData Then Exit For
    Next currentRow

    VerifyRecordsAreComplete = (Not missingData)
End Function

Private Function LoadTemplate() As String
    Const defaultTemplateFilename As String = "Speaking Evaluation Template.docx"
    Dim msgToDisplay As String, msgTitle As String
    Dim selectedPath As String, initialPath As String
    Dim basePath As String
    
    msgToDisplay = "Select where you have saved " & defaultTemplateFilename
    msgTitle = "Notice"
    
    initialPath = ThisWorkbook.Path
    ConvertOneDriveToLocalPath initialPath
    basePath = initialPath
    initialPath = initialPath & Application.PathSeparator & defaultTemplateFilename
    
    ' Verify 'Speaking Evaluation Template.docx' exists in the same folder as this Excel file
    selectedPath = IIf(Dir(initialPath) <> "", initialPath, "")
    
    ' Allow the user to select where it is saved if not found
    If selectedPath = "" Then
        ' This is only supported on MacOS if the user elects to install SpeakingEvals.scpt
        #If Mac Then
            If isAppleScriptInstalled Then
                MsgBox msgToDisplay, vbInformation, msgTitle
                selectedPath = AppleScriptTask("SpeakingEvals.scpt", "OpenTemplate", basePath)
            End If
        #Else
            Dim loadFileDialog As FileDialog
            Set loadFileDialog = Application.FileDialog(msoFileDialogFilePicker)
            
            MsgBox msgToDisplay, vbInformation, msgTitle
            
            With loadFileDialog
                .Title = "Load the Speaking Evaluation Template"
                .Filters.Clear
                .Filters.Add "Word Documents", "*.docx"
                .AllowMultiSelect = False
                .InitialFileName = initialPath
                
                If .Show = -1 Then
                    selectedPath = .SelectedItems(1)
                    ConvertOneDriveToLocalPath selectedPath
                Else
                    selectedPath = ""
                End If
            End With
            
            Set loadFileDialog = Nothing
        #End If
    End If
    
    If selectedPath = "" Then
        msgToDisplay = "No template was found. Process canceled."
        msgTitle = "Template Not Found"
        MsgBox msgToDisplay, vbExclamation, msgTitle
        LoadTemplate = ""
        
        If printDebugMessages Then
            Debug.Print "No template was found"
        End If
    ElseIf printDebugMessages Then
        Debug.Print "Template loaded: " & selectedPath
    End If
    
    LoadTemplate = selectedPath
End Function

Private Function SetSaveLocation(ByRef ws As Object) As String
    Dim msgToDisplay As String, msgTitle As String
    Dim selectedPath As String, pdfPath As String
    Dim userChoice As Integer
    Dim requestPermission As String
    
    selectedPath = ""
    userChoice = vbYes ' Default to yes in case a system doesn't support choosing a custom location
    
    msgToDisplay = "Would you like to save the reports in the same location as this Excel file?."
    msgTitle = "Notice"
    
    #If Mac Then
        ' This is only supported on MacOS if the user elects to install SpeakingEvals.scpt
        If isAppleScriptInstalled Then
            userChoice = MsgBox(msgToDisplay, vbYesNo, msgTitle)
        End If
    #Else
        userChoice = MsgBox(msgToDisplay, vbYesNo, msgTitle)
    #End If
    
    If userChoice = vbNo Then
        selectedPath = SetCustomLocation()
    Else
        ' The default directory's name will match the 'Class Days' value
        selectedPath = ThisWorkbook.Path & Application.PathSeparator & ws.Cells(4, 2).Value & Application.PathSeparator
        ConvertOneDriveToLocalPath selectedPath
        
        If printDebugMessages Then
            Debug.Print "Using default save path: " & vbNewLine & _
                        "Doc: " & selectedPath
        End If
        
        ' Check if the directory exists and create it if not found
        If Dir(selectedPath, vbDirectory) = "" Then
            If printDebugMessages Then
                Debug.Print "Path not found. Attempting to create."
            End If
            
            On Error Resume Next
            MkDir selectedPath
            On Error GoTo 0
        End If
        
        ' Sort out checking if a directory exists on MaOS. Perhaps an AppleScript would be best?
        'If Dir(selectedPath, vbDirectory) = "" Then
        '    If printDebugMessages Then
        '       Debug.Print "Error creating directories. Please select one manually."
        '    End If
            
        '    msgToDisplay = "Unable to create default folders. Please select where you would like to save to files."
        '    msgTitle = "Notice"
        '    MsgBox msgToDisplay, vbExclamation, msgTitle
        '
        '    selectedPath = SetCustomLocation()
        'Else
        '    If printDebugMessages Then
        '        Debug.Print "Directories successfully created. Continuing process."
        '    End If
        'End If
        ConvertOneDriveToLocalPath selectedPath
    End If
    
    SetSaveLocation = selectedPath
End Function

Private Function SetCustomLocation() As String
    Dim msgToDisplay As String, msgTitle As String
    Dim selectedPath As String, initialPath As String
    
    #If Mac Then
        selectedPath = AppleScriptTask("SpeakingEvals.scpt", "SelectSavePath", ThisWorkbook.Path & Application.PathSeparator)
    #Else
        Dim saveFolderDialog As FileDialog
        
        initialPath = ThisWorkbook.Path
        ConvertOneDriveToLocalPath initialPath
        
        If Right(initialPath, 1) <> Application.PathSeparator Then
            initialPath = initialPath & Application.PathSeparator
        End If
        
        Set saveFolderDialog = Application.FileDialog(msoFileDialogFolderPicker)
        
        With saveFolderDialog
            .Title = "Select Where to Save the Speaking Evaluations"
            .AllowMultiSelect = False
            .InitialFileName = initialPath
            
            If .Show = -1 Then
                selectedPath = .SelectedItems(1)
            Else
                selectedPath = ""
            End If
        End With
        
        Set saveFolderDialog = Nothing
    #End If
    
    ConvertOneDriveToLocalPath selectedPath
    
    If selectedPath <> "" And Right(selectedPath, 1) <> Application.PathSeparator Then
        selectedPath = selectedPath & Application.PathSeparator
    End If
    
    If selectedPath <> "" Then
        If printDebugMessages Then
            Debug.Print "Save location: " & selectedPath
        End If
    Else
        msgToDisplay = "No save folder was selected. Process canceled."
        msgTitle = "Save Location Not Found"
        MsgBox msgToDisplay, vbExclamation, msgTitle
    End If
    
    SetCustomLocation = selectedPath
End Function

Private Sub WriteReport(ByRef ws As Object, ByRef wordApp As Object, ByRef wordDoc As Object, ByVal currentRow As Integer, ByVal savePath As String, ByRef fileName As String, ByRef saveResult As Boolean)
    Dim nativeTeacher As String, koreanTeacher As String, classLevel As String, classTime As String, evalDate As String
    Dim englishName As String, koreanName As String, grammarScore As String, pronunciationScore As String, fluencyScore As String
    Dim mannerScore As String, contentScore As String, effortScore As String, commentText As String, overallGrade As String
    Dim signatureAdded As Boolean
    
    ' Data applicable to all reports
    nativeTeacher = ws.Cells(1, 2).Value
    koreanTeacher = ws.Cells(2, 2).Value
    classLevel = ws.Cells(3, 2).Value
    classTime = ws.Cells(4, 2).Value & "-" & ws.Cells(5, 2).Value
    evalDate = ws.Cells(6, 2).Value
    evalDate = Format(Date, "MMM. YYYY")
    
    ' Data specific to each student
    englishName = ws.Cells(currentRow, 1).Value
    koreanName = ws.Cells(currentRow, 2).Value
    grammarScore = ws.Cells(currentRow, 3).Value
    pronunciationScore = ws.Cells(currentRow, 4).Value
    fluencyScore = ws.Cells(currentRow, 5).Value
    mannerScore = ws.Cells(currentRow, 6).Value
    contentScore = ws.Cells(currentRow, 7).Value
    effortScore = ws.Cells(currentRow, 8).Value
    commentText = ws.Cells(currentRow, 9).Value
    overallGrade = CalculateOverallGrade(ws, currentRow)
    
    fileName = koreanTeacher & "(" & classTime & ")" & " - " & koreanName & "(" & englishName & ")"
    
    If printDebugMessages Then
        Debug.Print "Saving to: " & savePath & vbNewLine & _
                    "Saving as: " & fileName
    End If
    
    ' Add code to explicitly set the border type
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
    
    If signatureAdded = False Then
        InsertSignature wordDoc
    End If
    
    On Error Resume Next
    #If Mac Then
        ' The export feature is a bit flaky on MacOS, so we need to do a full SaveAs2. Only results in a minimal time loss.
        wordDoc.SaveAs2 fileName:=(savePath & fileName & ".pdf"), FileFormat:=17, AddtoRecentFiles:=False, EmbedTrueTypeFonts:=True
    #Else
        wordDoc.ExportAsFixedFormat OutputFileName:=(savePath & fileName & ".pdf"), ExportFormat:=17, BitmapMissingFonts:=True
    #End If
    
    saveResult = (Err.Number = 0)
    
    If printDebugMessages Then
        If saveResult Then
            Debug.Print "Successfully saved: " & fileName
        Else
            Debug.Print "Failed to save." & _
                        "Error Number: " & Err.Number & vbNewLine & _
                        "Error Description: " & Err.Description
            Err.Clear
        End If
    End If
    On Error GoTo 0
End Sub

Private Function CalculateOverallGrade(ByRef ws As Worksheet, ByVal currentRow As Integer) As String
    Dim scoreRange As Range, gradeCell As Range
    Dim totalScore As Integer, avgScore As Integer, numericScore As Integer
    
    totalScore = 0
    Set scoreRange = ws.Range("C" & currentRow & ":H" & currentRow)
    
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
    
    avgScore = Int(totalScore / 6)
    
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
    
    Const signatureShapeName As String = "mySignature"
    ' These numbers make no sense, but they work.
    Const absoluteLeft As Double = 332.4
    Const absoluteTop As Double = 684
    Const maxWidth As Double = 144
    Const maxHeight As Double = 40
    
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
            SaveSignature signatureShapeName, newImagePath
        ElseIf Dir(signaturePath & "mySignature.png") <> "" Then
            newImagePath = signaturePath & "mySignature.png"
        ElseIf Dir(signaturePath & "mySignature.jpg") <> "" Then
            newImagePath = signaturePath & "mySignature.jpg"
        Else
            Exit Sub
        End If
    End If
    
    Set newImageShape = wordDoc.Shapes.AddPicture(fileName:=newImagePath, LinkToFile:=False, SaveWithDocument:=True)
    newImageShape.Name = "Signature"
    
    ' Maintain the aspect ratio and resize if needed
    aspectRatio = newImageShape.Width / newImageShape.Height
    If maxWidth / maxHeight > aspectRatio Then
        ' Adjust width to fit within max height
        imageWidth = maxHeight * aspectRatio
        imageHeight = maxHeight
    Else
        ' Adjust height to fit within max width
        imageWidth = maxWidth
        imageHeight = maxWidth / aspectRatio
    End If

    ' Position and resize the image
    With newImageShape
        .LockAspectRatio = msoTrue
        .Left = absoluteLeft
        .Top = absoluteTop
        .Width = imageWidth
        .Height = imageHeight
        
        ' Ensure positioning relative to the page
        .RelativeHorizontalPosition = 1
        .RelativeVerticalPosition = 1
    End With
End Sub

Private Sub SaveSignature(ByVal signatureShapeName As String, ByRef savePath As String)
    Dim signSheet As Worksheet, tempSheet As Worksheet
    Dim signatureshp As Shape, chrt As ChartObject
    
    Set signatureshp = ThisWorkbook.Worksheets("Instructions").Shapes(signatureShapeName)
    
    #If Mac Then
        savePath = Environ("TMPDIR") & "tempSignature.png"
    #Else
        savePath = Environ("TEMP") & Application.PathSeparator & "tempSignature.png"
    #End If
    
    ConvertOneDriveToLocalPath savePath
    
    Sheets.Add(, Sheets(Sheets.count)).Name = "Temp_signature"
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

Private Sub WaitTimer(ByVal timeToWait As Integer)
    Dim startTime As Single: startTime = Timer
    
    Do While Timer < startTime + timeToWait
        DoEvents
    Loop
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
            If printDebugMessages Then
                Debug.Print "Missing shape: " & shapeNames(i)
            End If
            msgToDisplay = "There is a critical error with the template. Please redownload a copy of the original and try again."
            MsgBox msgToDisplay, vbExclamation, "Error!"
            VerifyAllDocShapesExist = False
            Exit Function
        End If
        On Error GoTo 0
    Next i
                       
    VerifyAllDocShapesExist = True
End Function

Private Function WordDocShapeExists(ByRef wordDoc As Object, ByVal shapeName As String) As Boolean
    Dim shp As Object, grpItem As Object
    
    If printDebugMessages Then
        Debug.Print "Search for shape: " & shapeName
    End If
    
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
    
    If printDebugMessages Then
        Debug.Print "Unable to find shape: " & shapeName
    End If
    
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
    
    ' While mostly invisible to the user, cloud storage services like iCloud and OneDrive actually change
    ' where files are stored. This is especially problematic with OneDrive, as it uses URIs (internet links)
    ' instead of normal file paths. This examines the path of the file on the user's computer and converts it
    ' back into a local path so that files can be opened and saved properly.
    
    If Left(selectedPath, 23) = "https://d.docs.live.net" Or Left(selectedPath, 11) = "OneDrive://" Then
        For i = 1 To 4 ' Everything befor the 4th '/' is the OneDrive URI and needs to be removed.
            selectedPath = Mid(selectedPath, InStr(selectedPath, "/") + 1)
        Next
        
        ' Append the local file directory to the beginning of the trimmed 'selectedPath' above.
        #If Mac Then
            selectedPath = "/Users/" & Environ("USER") & "/Library/CloudStorage/OneDrive-Personal/" & selectedPath
        #Else
            selectedPath = Replace(selectedPath, "/", "\")
            selectedPath = Environ$("OneDrive") & "\" & selectedPath
        #End If
    Else
        ' This may not be needed, but is here just in case. After some more testing, this may either be expanded or removed.
        #If Mac Then
            If InStr(1, selectedPath, "iCloud Drive", vbTextCompare) > 0 Then
                For i = 1 To 4 ' Strip away the iCloud part of the filepath (everything before the 6th '/')
                    selectedPath = Mid(selectedPath, InStr(selectedPath, "/") + 1)
                Next
                
                If printDebugMessages Then
                    Debug.Print "Trimmed iCloud file path: " & selectedPath
                End If
                
                selectedPath = "/Users/" & Environ("USER") & "/Library/Mobile Documents/com~apple~CloudDocs/" & selectedPath
            End If
        #End If
    End If
End Sub

#If Mac Then
Private Function checkForAppleScript() As Boolean
    Dim appleScriptPath As String, returnedPath As String
    
    appleScriptPath = "/Users/" & Environ("USER") & "/Library/Application Scripts/com.microsoft.Excel/SpeakingEvals.scpt"
    
    If printDebugMessages Then
        Debug.Print "Locating SpeakingEvals.scpt."
    End If
    
    On Error Resume Next
    returnedPath = Dir(appleScriptPath, vbDirectory)
    
    If returnedPath <> vbNullString Then
        If printDebugMessages Then
            Debug.Print "Successfully found at: " & appleScriptPath
        End If
        On Error GoTo 0
        checkForAppleScript = True
    End If
    
    If returnedPath = vbNullString Then
        If printDebugMessages Then
            Debug.Print "Not found!" & returnedPath
        End If
        On Error GoTo 0
        checkForAppleScript = False
    End If
End Function

Private Sub PromptToInstallAppleScript()
    Dim msgToDisplay As String, msgTitle As String, userChoice As Integer
    Dim curlCommand As String
    
    msgToDisplay = "The SpeakingEvals.scpt file is not currently installed. It is optional, but having it installed enables a few additional features to make usage on Macs smoother. Would you like to download and install it?"
    msgTitle = "Missing Data!"
    userChoice = MsgBox(msgToDisplay, vbYesNo, msgTitle)
    
    If userChoice = vbYes Then
        curlCommand = "curl -L -o ~/Library/Application\ Scripts/com.microsoft.Excel/SpeakingEvals.scpt https://github.com/papercutter0324/SpeakingEvals/raw/main/SpeakingEvals.scpt"
        On Error Resume Next
        MacScript "do shell script """ & curlCommand & """"
        On Error GoTo 0
        
        If Not checkForAppleScript Then
            msgToDisplay = "There was an error installing the file automatically. Please have a look at the Instructions sheet for how to easily install it manually."
            msgTitle = "Error!"
            MsgBox msgToDisplay, vbExclamation, msgTitle
        End If
    End If
    
    isAppleScriptInstalled = checkForAppleScript()
End Sub

Private Function GetLocalOneDrivePath(ByVal destinationPath As String) As String
    GetLocalOneDrivePath = Replace(MacScript("return POSIX path of (path to desktop folder) as string"), "/Desktop", "/Library/CloudStorage/OneDrive-Personal/Desktop") & destinationPath
End Function

Private Sub RequestMacExcelPermissions(ByVal templatePath As String, ByVal savePath As String)
    Dim currentPath As String, tempSignaturePath As String
    Dim fileAccessGranted As Boolean
    Dim filePermissionCandidates As Variant
    
    currentPath = ThisWorkbook.Path
    tempSignaturePath = Environ("TMPDIR") & "tempSignature.png"
    
    ConvertOneDriveToLocalPath currentPath
    ConvertOneDriveToLocalPath tempSignaturePath
    
    filePermissionCandidates = Array(currentPath, savePath, templatePath, tempSignaturePath)

    fileAccessGranted = GrantAccessToMultipleFiles(filePermissionCandidates)
    
    ' RequestMacWordPermissions currentPath, templatePath, savePath, tempSignaturePath
End Sub

Private Sub RequestMacWordPermissions(ByVal currentPath As String, ByVal templatePath As String, ByVal savePath As String, ByVal tempSignaturePath As String)
    Dim requestScript As String
    
    ' Try to get Word to ask for all permissions at the same time in order to cut down on user prompts
    requestScript = "tell application ""Microsoft Word""" & vbCrLf & _
                    "   set filePaths to {""" & templatePath & """, """ & savePath & """, """ & tempSignaturePath & """}" & vbCrLf & _
                    "   repeat with filePath in filePaths" & vbCrLf & _
                    "       set myFile to POSIX file filePath" & vbCrLf & _
                    "       open myFile" & vbCrLf & _
                    "       close myFile" & vbCrLf & _
                    "   end repeat" & vbCrLf & _
                    "end tell"
                    
    On Error Resume Next
    MacScript requestScript
    If Err.Number <> 0 Then
        MsgBox "Error requesting Word permissions: " & Err.Description, vbCritical
    End If
    On Error GoTo 0
End Sub
#End If

Private Function VerifyFileOrFolderExists(ByVal pathToCheck As String) As Boolean
    ' Borrowed from my Angry Birds Trivia.pptm game. Will update as it becomes needed.
    #If Mac Then
        Dim pathExists As Boolean
        pathExists = AppleScriptTask("AngryBirds.scpt", "ExistsFile", pathToCheck)
        
        If Not pathExists Then
            pathExists = AppleScriptTask("AngryBirds.scpt", "ExistsFolder", pathToCheck)
        End If
        
        VerifyFileOrFolderExists = pathExists
    #Else
        Dim fs As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
        
        VerifyFileOrFolderExists = (fs.fileExists(pathToCheck) Or fs.FolderExists(pathToCheck))
    #End If
End Function
