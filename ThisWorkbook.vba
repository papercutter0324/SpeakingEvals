Option Explicit
#Const PRINT_DEBUG_MESSAGES = True
Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
Const APPLE_SCRIPT_SPLIT_KEY = "-,-"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  Auto-run Sub on Startup and Worksheet Switching
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Workbook_Open()
    Const CURL_COMMAND_TEXT As String = "curl -L -o ~/Library/Application\ Scripts/com.microsoft.Excel/SpeakingEvals.scpt https://github.com/papercutter0324/SpeakingEvals/raw/main/SpeakingEvals.scpt"
    Dim ws As Worksheet, shps As Shapes
    Dim startupMessageToDisplay As String
    
    ThisWorkbook.Sheets("Instructions").Activate
    startupMessageToDisplay = "Initial"
    DisplayStartupMessage startupMessageToDisplay
    
    On Error GoTo ReenableEvents
    Application.EnableEvents = False
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect
        
        Select Case ws.Name
            Case "Instructions"
                SetLayoutInstructions ThisWorkbook, ws
            Case "MacOS Users"
                Set shps = ws.Shapes
                shps("cURL_Command").TextFrame2.TextRange.Characters.Text = CURL_COMMAND_TEXT
                
                #If Mac Then
                    Dim scriptResult As Boolean
                    scriptResult = AreAppleScriptsInstalled
                #Else
                    shps("Button_SpeakingEvalsScpt_Missing").Visible = True
                    shps("Button_DialogToolkit_Missing").Visible = True
                    shps("Button_EnhancedDialogs_Disable").Visible = True
                    shps("Button_SpeakingEvalsScpt_Installed").Visible = False
                    shps("Button_DialogToolkit_Installed").Visible = False
                    shps("Button_EnhancedDialogs_Enable").Visible = False
                #End If
                
                Set shps = Nothing
                SetLayoutMacOSUsers ThisWorkbook
            Case "mySignature"
                SetLayoutMySignature ThisWorkbook
            Case Else
                AutoPopulateEvaluationDateValues ws
                SetLayoutClassRecords ThisWorkbook, ws
        End Select
        
        ws.Protect
        ws.EnableSelection = xlUnlockedCells
    Next ws
    
    startupMessageToDisplay = "Complete"
    
    #If Mac Then
        If Not scriptResult Then
            startupMessageToDisplay = "SpeakEvals.scpt Reminder"
            ThisWorkbook.Sheets("MacOS Users").Activate
        End If
    #End If
        
    DisplayStartupMessage startupMessageToDisplay

ReenableEvents:
    Application.EnableEvents = True
End Sub

Private Sub Workbook_SheetActivate(ByVal ws As Object)
    Const CURL_COMMAND_TEXT As String = "curl -L -o ~/Library/Application\ Scripts/com.microsoft.Excel/SpeakingEvals.scpt https://github.com/papercutter0324/SpeakingEvals/raw/main/SpeakingEvals.scpt"
    
    Application.EnableEvents = False
    
    Select Case ws.Name
        Case "Instructions"
            SetLayoutInstructions ThisWorkbook, ws
        Case "MacOS Users"
            ws.Shapes("cURL_Command").TextFrame2.TextRange.Characters.Text = CURL_COMMAND_TEXT
            SetLayoutMacOSUsers ThisWorkbook
        Case "mySignature"
            SetLayoutMySignature ThisWorkbook
        Case Else
            SetLayoutClassRecords ThisWorkbook, ws
    End Select
    
    Application.EnableEvents = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    #If Mac Then
        Dim resourcesFolder As String
        
        resourcesFolder = ThisWorkbook.Path & "/Resources"
        ConvertOneDriveToLocalPath resourcesFolder
        RemoveDialogToolKit resourcesFolder
    #End If
End Sub

Private Sub DisplayStartupMessage(ByVal startupStage As String)
    Dim msgToDisplay As String, dialogSize As Integer, msgResult As Integer
    
    Select Case startupStage
        Case "Initial"
            msgToDisplay = "Please wait a few moments while a self-check is performed to ensure all sheets and cells are properly formatted. " & _
                           "All existing information will br preserved, and another message will appear to inform you the process is complete."
            dialogSize = 430
        Case "Complete"
            msgToDisplay = "Process complete!"
            dialogSize = 120
        Case "SpeakEvals.scpt Reminder"
            msgToDisplay = "You must install SpeakingEvals.scpt in order for this file to fuction properly. Please follow these instructions and " & _
                           "read the notices about why you will see popups for System Events and file & folder permission requests."
            dialogSize = 470
    End Select
    
    msgResult = DisplayMessage(msgToDisplay, vbInformation, "Welcome!", dialogSize)
End Sub

Private Sub AutoPopulateEvaluationDateValues(ByRef ws As Worksheet)
    Dim dateValue As Range
    Dim messageText As String, msgResult As Variant
    
    On Error Resume Next
    Application.EnableEvents = False
    
    If ws.Range("A6").Value = "Evaluation Date:" Then
        Set dateValue = ws.Range("C6")

        If IsEmpty(dateValue) Then
            ' Add date if not entered yet
            dateValue.Value = Format(Date, "MMM. YYYY")
        ElseIf IsDate(dateValue.Value) And dateValue.Value < Date - 21 Then
            ' Update the date if the current value is over 21 days ago
            dateValue.Value = Format(Date, "MMM. YYYY")
        ElseIf Not IsDate(dateValue.Value) Then
            ' Display an error message if an invalid date is found
            messageText = "An invalid date has been found on worksheet " & ws.Name & "." & vbNewLine & "Please enter a valid date."
            msgResult = DisplayMessage(messageText, vbInformation, "Invalid Date!")
        End If
    End If
    
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Private Sub SetLayoutInstructions(ByRef wb As Workbook, ByRef ws As Worksheet)
    Dim shp As Shapes, refShape As Shape, currentshp As Shape
    Dim btnHeight As Double, btnWidth As Double, cellSpacing As Double
    Dim btnNamesArray As Variant
    Dim i As Integer
    
    Set shp = ws.Shapes
    
    shp("Title Bar - Instructions").Top = 15
    shp("Title Bar - Instructions").Left = 15
    shp("Title Bar - Instructions").Height = 58
    shp("Title Bar - Instructions").Width = 1165
    Set refShape = shp("Title Bar - Instructions")
    
    shp("Message - Instructions").Top = refShape.Top + refShape.Height
    shp("Message - Instructions").Left = refShape.Left
    shp("Message - Instructions").Height = 610
    shp("Message - Instructions").Width = refShape.Width
    Set refShape = shp("Message - Instructions")
    
    shp("Title Bar - Seeing the Code").Top = refShape.Top + refShape.Height + 20
    shp("Title Bar - Seeing the Code").Left = refShape.Left
    shp("Title Bar - Seeing the Code").Height = shp("Title Bar - Instructions").Height
    shp("Title Bar - Seeing the Code").Width = refShape.Width
    Set refShape = shp("Title Bar - Seeing the Code")
    
    shp("Message - Seeing the Code").Top = refShape.Top + refShape.Height
    shp("Message - Seeing the Code").Left = refShape.Left
    shp("Message - Seeing the Code").Height = 265
    shp("Message - Seeing the Code").Width = refShape.Width
    
    Set refShape = shp("Title Bar - Instructions")
    
    shp("Title Bar - Warning").Top = refShape.Top
    shp("Title Bar - Warning").Left = refShape.Left + refShape.Width + 20
    shp("Title Bar - Warning").Height = refShape.Height
    shp("Title Bar - Warning").Width = 340
    Set refShape = shp("Title Bar - Warning")
    
    shp("Message - Warning").Top = refShape.Top + refShape.Height
    shp("Message - Warning").Left = refShape.Left
    shp("Message - Warning").Height = 150
    shp("Message - Warning").Width = refShape.Width
    Set refShape = shp("Message - Warning")
    
    shp("Title Bar - Important Files").Top = refShape.Top + refShape.Height + 20
    shp("Title Bar - Important Files").Left = refShape.Left
    shp("Title Bar - Important Files").Height = shp("Title Bar - Instructions").Height
    shp("Title Bar - Important Files").Width = refShape.Width
    Set refShape = shp("Title Bar - Important Files")
    
    shp("Message - Important Files").Top = refShape.Top + refShape.Height
    shp("Message - Important Files").Left = refShape.Left
    shp("Message - Important Files").Height = 725
    shp("Message - Important Files").Width = refShape.Width
    
    Set refShape = shp("Message - Seeing the Code")
    btnHeight = 70
    btnWidth = 200
    cellSpacing = (refShape.Width - 40 - (btnWidth * 5)) / 4
    btnNamesArray = Array("Button_Speadsheet", "Button_Font", "Button_ReportTemplate", "Button_SignatureTemplate", "Button_SourceCode")
    
    For i = LBound(btnNamesArray) To UBound(btnNamesArray)
        Set currentshp = shp(btnNamesArray(i))
        currentshp.Height = btnHeight
        currentshp.Width = btnWidth
        currentshp.Left = refShape.Left + 20 + i * (btnWidth + cellSpacing)
        currentshp.Top = refShape.Top + refShape.Height - btnHeight - 20
    Next i
End Sub

Private Sub SetLayoutMacOSUsers(ByRef wb As Workbook)
    Dim refShape As Shape, currentBtn As Shape
    Dim buttonNamesArray As Variant, i As Integer
    
    buttonNamesArray = Array("Button_SpeakingEvalsScpt_Installed", "Button_SpeakingEvalsScpt_Missing", "Button_DialogToolkit_Installed", _
                            "Button_DialogToolkit_Missing", "Button_EnhancedDialogs_Enable", "Button_EnhancedDialogs_Disable")
    
    With wb.Sheets("MacOS Users")
        .Unprotect
        With .Shapes
            With .Item("Title Bar")
                .Height = 58
                .Width = 1285
                .Top = 15
                .Left = 15
            End With
            
            Set refShape = .Item("Title Bar")
            With .Item("Message")
                .Height = 680
                .Width = refShape.Width
                .Top = refShape.Top + refShape.Height
                .Left = refShape.Left
            End With
                    
            Set refShape = .Item("Message")
            With .Item("cURL_Command")
                .Height = 115
                .Width = 560
                .Top = refShape.Top
                .Left = refShape.Left + refShape.Width - .Width
            End With
            
            For i = LBound(buttonNamesArray) To UBound(buttonNamesArray)
                Set currentBtn = .Item(buttonNamesArray(i))
                currentBtn.Height = 70
                currentBtn.Width = 200
                currentBtn.Top = refShape.Top + refShape.Height - currentBtn.Height - 20
                
                Select Case buttonNamesArray(i)
                    Case "Button_SpeakingEvalsScpt_Installed", "Button_SpeakingEvalsScpt_Missing"
                        currentBtn.Left = refShape.Left + 70
                    Case "Button_DialogToolkit_Installed", "Button_DialogToolkit_Missing"
                        currentBtn.Left = refShape.Left + 350
                    Case "Button_EnhancedDialogs_Enable", "Button_EnhancedDialogs_Disable"
                        currentBtn.Left = refShape.Left + 630
                End Select
            Next i
        End With
        .Protect
    End With
End Sub

Private Sub SetLayoutMySignature(ByRef wb As Workbook)
    Dim refShape As Shape
    
    With wb.Sheets("mySignature")
        .Unprotect
        With .Shapes
            With .Item("Title Bar")
                .Height = 58
                .Width = 1270
                .Top = 15
                .Left = 15
            End With
            
            Set refShape = .Item("Title Bar")
            With .Item("Message")
                .Height = 840
                .Width = refShape.Width
                .Top = refShape.Top + refShape.Height
                .Left = refShape.Left
            End With
                    
            With .Item("Signature Title Bar")
                .Height = refShape.Height
                .Width = 300
                .Top = refShape.Top
                .Left = refShape.Left + refShape.Width - .Width
            End With
            
            Set refShape = .Item("Signature Title Bar")
            With .Item("Signature Container")
                .Height = 130
                .Width = refShape.Width
                .Top = refShape.Top + refShape.Height
                .Left = refShape.Left
            End With
            
            Set refShape = .Item("Signature Container")
            On Error Resume Next
            If DoesShapeExist(wb.Sheets("mySignature"), "mySignature_Placeholder") Then
                With .Item("mySignature_Placeholder")
                    .Width = 378.09
                    .Height = 105
                    .Top = refShape.Top + (refShape.Height / 2) - (.Height / 2)
                    .Left = refShape.Left + (refShape.Width / 2) - (.Width / 2)
                End With
            End If
            
            If DoesShapeExist(wb.Sheets("mySignature"), "mySignature") Then
                With .Item("mySignature")
                    .Top = refShape.Top + (refShape.Height / 2) - (.Height / 2)
                    .Left = refShape.Left + (refShape.Width / 2) - (.Width / 2)
                End With
            End If
            On Error GoTo 0
        End With
        .Protect
    End With
End Sub

Private Function DoesShapeExist(ByVal ws As Worksheet, ByVal shapeName As String) As Boolean
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    DoesShapeExist = Not shp Is Nothing
    On Error GoTo 0
End Function

Private Sub SetLayoutClassRecords(ByRef wb As Workbook, ByRef ws As Worksheet)
    Dim btnGenerateProofs As Shape, btnGenerateReports As Shape, btnOpenTypingSite As Shape, btnRepairLayout As Shape
    Dim shadingRanges As Variant, borderRanges As Variant, borderTypes As Variant, borderSetting As Variant
    Dim cellTop As Double, cellHeight As Double, cellLeft As Double, cellWidth As Double
    Dim cellVerticalSpacing As Double, cellHorizontalSpacing As Double
    Dim buttonHeight As Double, buttonWidth As Double
    Dim mainBorders As Range, dashedBorders As Range, noInsideBorders As Range
    Dim unlockedCells As Range, lockedCells As Range, currentRng As Range
    Dim i As Integer
        
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ws.Unprotect
    
    shadingRanges = Array( _
        Array("classInfoShadingRange", "A1:C6", RGB(255, 255, 255)), _
        Array("indexEngKorHeaderShapingRange", "A7:C7", RGB(197, 217, 241)), _
        Array("indexNumberShadingRange", "A8:A32", RGB(217, 217, 217)), _
        Array("EngKorNameShadingRange", "B8:C32", RGB(255, 255, 255)), _
        Array("grammarHeaderShadingRange", "D1:D7", RGB(177, 160, 199)), _
        Array("grammarValuesShadingRange", "D8:D32", RGB(228, 223, 236)), _
        Array("pronunciationHeaderShadingRange", "E1:E7", RGB(146, 205, 220)), _
        Array("pronunciationValuesShadingRange", "E8:E32", RGB(218, 238, 243)), _
        Array("fluencyHeaderShadingRange", "F1:F7", RGB(218, 150, 148)), _
        Array("fluencyValuesShadingRange", "F8:F32", RGB(242, 220, 219)), _
        Array("mannerHeaderShadingRange", "G1:G7", RGB(250, 191, 143)), _
        Array("mannerValuesShadingRange", "G8:G32", RGB(253, 233, 217)), _
        Array("contentHeaderShadingRange", "H1:H7", RGB(196, 215, 155)), _
        Array("contentValuesShadingRange", "H8:H32", RGB(235, 241, 222)), _
        Array("overallEffortHeaderShadingRange", "I1:I7", RGB(149, 179, 215)), _
        Array("overallEffortValuesShadingRange", "I8:I32", RGB(220, 230, 241)), _
        Array("buttonShadingRange", "J1:J5", RGB(255, 255, 255)), _
        Array("commentHeaderShadingRange", "J6:J7", RGB(191, 191, 191)), _
        Array("commentValuesShadingRange", "J8:J32", RGB(242, 242, 242)), _
        Array("notesHeaderShadingRange", "K1:K7", RGB(196, 189, 151)), _
        Array("notesValuesShadingRange", "K8:K32", RGB(221, 217, 196)) _
    )
    
    For i = LBound(shadingRanges) To UBound(shadingRanges)
        Set currentRng = ws.Range(shadingRanges(i)(1))
        On Error Resume Next
        If Not wb.Names(shadingRanges(i)(0)) Is Nothing Then wb.Names(shadingRanges(i)(0)).Delete
        wb.Names.Add Name:=shadingRanges(i)(0), RefersTo:=currentRng
        currentRng.Interior.Color = shadingRanges(i)(2)
        VerifyTipMessages ws, shadingRanges(i)(0), currentRng
        On Error GoTo 0
    Next i
    
    With ws
        .Columns("A").ColumnWidth = 7
        .Columns("B:C").ColumnWidth = 18
        .Columns("D:I").ColumnWidth = 21
        .Columns("J").ColumnWidth = 102.5
        .Columns("K").ColumnWidth = 44.17
        .Rows("1:6").RowHeight = 30
        .Rows("7").RowHeight = 25
        .Rows("8:32").RowHeight = 50
    End With
        
    cellTop = ws.Cells(1, 10).Top
    cellHeight = ws.Cells(1, 10).Height * 5
    cellLeft = ws.Cells(1, 10).Left
    cellWidth = ws.Cells(1, 10).Width
    buttonHeight = 50
    buttonWidth = 187
    cellVerticalSpacing = (cellHeight - (2 * buttonHeight)) / 3
    cellHorizontalSpacing = (cellWidth - (2 * buttonWidth)) / 3
        
    Set btnGenerateProofs = ws.Shapes("Button_GenerateProofs")
    Set btnGenerateReports = ws.Shapes("Button_GenerateReports")
    Set btnOpenTypingSite = ws.Shapes("Button_OpenTypingSite")
    Set btnRepairLayout = ws.Shapes("Button_RepairLayout")
        
    btnGenerateProofs.Height = buttonHeight
    btnGenerateProofs.Width = buttonWidth
    btnGenerateProofs.Top = cellTop + cellVerticalSpacing
    btnGenerateProofs.Left = cellLeft + cellHorizontalSpacing
        
    btnGenerateReports.Height = buttonHeight
    btnGenerateReports.Width = buttonWidth
    btnGenerateReports.Top = cellTop + (cellVerticalSpacing * 2) + buttonHeight
    btnGenerateReports.Left = cellLeft + cellHorizontalSpacing

    btnOpenTypingSite.Height = buttonHeight
    btnOpenTypingSite.Width = buttonWidth
    btnOpenTypingSite.Top = cellTop + cellVerticalSpacing
    btnOpenTypingSite.Left = cellLeft + (cellHorizontalSpacing * 2) + buttonWidth

    btnRepairLayout.Height = buttonHeight
    btnRepairLayout.Width = buttonWidth
    btnRepairLayout.Top = cellTop + (cellVerticalSpacing * 2) + buttonHeight
    btnRepairLayout.Left = cellLeft + (cellHorizontalSpacing * 2) + buttonWidth
        
    On Error Resume Next
    Set mainBorders = ws.Range("A1:C6,A8:A32,B8:C32,D8:I32,J8:J32,K8:K32")
    Set dashedBorders = ws.Range("A1:C6,A8:A32,B8:C32,D8:I32,J8:J32,K8:K32")
    Set noInsideBorders = ws.Range("D1:I6,J1:J5,K1:K6")
    Set unlockedCells = ws.Range("C1:C6,B8:K32")
    Set lockedCells = ws.Range("A1:B6,D1:K6,A7:K7,A8:A32")
    
    With mainBorders.Borders
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    
    With dashedBorders.Borders
        .Item(xlInsideHorizontal).LineStyle = xlDash
        .Item(xlInsideHorizontal).Weight = xlThin
        .Item(xlInsideVertical).LineStyle = xlLineStyleNone
    End With
    
    With noInsideBorders.Borders
        .Item(xlInsideHorizontal).LineStyle = xlLineStyleNone
        .Item(xlInsideVertical).LineStyle = xlLineStyleNone
    End With
    
    unlockedCells.Locked = False
    lockedCells.Locked = True
    On Error GoTo 0
    
    ws.Protect
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub VerifyTipMessages(ByRef ws As Worksheet, ByVal rangeType As String, ByRef cellRange As Range)
    Dim currentCell As Range, rangeCol As Range, rangeCell As Range
    Dim dateInputMessage As Variant, validationValues As Variant
    Dim validationInputTitle As String, validationInputMessage As String
    Dim columnFont As String, columnFontSize As Integer, columnFontBold As Boolean
    Dim i As Integer
    
    Select Case rangeType
        Case "classInfoShadingRange"
            Select Case Application.International(xlDateOrder)
               Case 0
                   dateInputMessage = "MM/DD/YYYY" & vbNewLine & "or MM/YYYY."
               Case 1
                   dateInputMessage = "DD/MM/YYYY" & vbNewLine & "or MM/YYYY."
               Case 2
                   dateInputMessage = "YYYY/MM/DD" & vbNewLine & "or MM/YYYY."
            End Select
            
            validationValues = Array( _
                Array("Native Teacher's Name", "Please enter just your" & vbNewLine & "name, no suffix or title" & vbNewLine & "like ""tr."""), _
                Array("Korean Teacher's Name", "Please write their Korean name. The parents are unlikely to know their English name."), _
                Array("Class Level", "Click on the down arrow and choose the class's level from the list."), _
                Array("Class Days", "Select the days when you see this class." & vbNewLine & vbNewLine & "For Athena and Song's classes, use Class-1 and Class-2 to help organize split classes."), _
                Array("Class Time", "Select what time you have Class 1 each week. Scroll to see more options." & vbNewLine & vbNewLine & "This is to help you keep track of which class this is; it won't appear on the final reports."), _
                Array("Date Format", dateInputMessage) _
            )
            
            For i = 1 To 6
                columnFont = IIf(i = 2, "Batang", "Calibri")
                SetCellFormating ws.Cells(i, 3), validationValues(i - 1)(0), validationValues(i - 1)(1), columnFont, 14, False
            Next i
        Case "EngKorNameShadingRange", "grammarValuesShadingRange", "pronunciationValuesShadingRange", "fluencyValuesShadingRange", "mannerValuesShadingRange", "contentValuesShadingRange", "overallEffortValuesShadingRange", "commentValuesShadingRange", "notesValuesShadingRange"
            columnFont = "Calibri"
            columnFontBold = False
            For Each rangeCol In cellRange.Columns
                Select Case rangeCol.Column
                    Case 2
                        validationInputTitle = "Character Limit"
                        validationInputMessage = "    30 characters"
                        columnFontSize = 18
                    Case 3
                        validationInputTitle = ""
                        validationInputMessage = ""
                        columnFont = "Batang"
                        columnFontSize = 20
                    Case 4 To 9
                        validationInputTitle = "Enter a Grade"
                        validationInputMessage = "Valid Letter Grades" & vbNewLine & "  A+ / A / B+ / B / C" & vbNewLine & vbNewLine & "Valid Numeric Scores   " & vbNewLine & "  1 ~ 5"
                        columnFontSize = 22
                        columnFontBold = True
                    Case 10
                        validationInputTitle = "Feedback Sandwich"
                        validationInputMessage = "FORMAT:" & vbNewLine & "  Good-Bad-Good" & vbNewLine & vbNewLine & "LIMIT:" & vbNewLine & "  315 characters"
                        columnFontSize = 14
                    Case 11
                        validationInputTitle = ""
                        validationInputMessage = ""
                        columnFontSize = 14
                End Select
                
                For Each rangeCell In rangeCol.Cells
                    SetCellFormating rangeCell, validationInputTitle, validationInputMessage, columnFont, columnFontSize, columnFontBold
                Next rangeCell
            Next rangeCol
    End Select
End Sub

Private Sub SetCellFormating(ByRef wsCell As Range, ByVal validationInputTitle As String, ByVal validationInputMessage As String, ByVal columnFont As String, ByVal columnFontSize As Integer, ByVal columnFontBold As Boolean)
    On Error Resume Next
    With wsCell
        With .Validation
            If .inputTitle <> validationInputTitle Then
                .Delete
                .Add Type:=xlValidateInputOnly, _
                     AlertStyle:=xlValidAlertStop
                .inputTitle = validationInputTitle
                .inputMessage = validationInputMessage
                .ShowError = False
            End If
        End With
        With .Font
            .Name = columnFont
            .Size = columnFontSize
            .Bold = columnFontBold
            .Italic = False
            .Underline = False
        End With
    End With
    On Error GoTo 0
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  Main Subs and Functions
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Main()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim clickedButtonName As String: clickedButtonName = Application.Caller
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Beginning tasks." & vbNewLine & _
                    "    Active Worksheet = " & ws.Name
    #End If
    
    #If Mac Then
        Dim msgToDiplay As String
        
        If Not AreAppleScriptsInstalled(True) Then
            RemindUserToInstallSpeakingEvalsScpt
            Exit Sub
        End If
    #End If
    
    On Error GoTo ReenableEvents
    Application.EnableEvents = False
    
    Select Case clickedButtonName
        #If Mac Then
        Case "Button_EnhancedDialogs_Enable", "Button_EnhancedDialogs_Disable"
            ToogleMacSettingsButtons ws, clickedButtonName
        #End If
        Case "Button_GenerateReports", "Button_GenerateProofs"
            GenerateReports ws, clickedButtonName
            ws.Activate ' Ensure the right worksheet is being shown when finished.
        Case "Button_RepairLayout"
            SetLayoutClassRecords ThisWorkbook, ws
    End Select
    
ReenableEvents:
    Application.EnableEvents = True
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
    
    Select Case roundedScore
        Case 5: CalculateOverallGrade = "A+"
        Case 4: CalculateOverallGrade = "A"
        Case 3: CalculateOverallGrade = "B+"
        Case 2: CalculateOverallGrade = "B"
        Case 1: CalculateOverallGrade = "C"
    End Select
End Function

Private Sub ExportSignatureFromExcel(ByVal SIGNATURE_SHAPE_NAME As String, ByRef savePath As String)
    Dim signSheet As Worksheet, tempSheet As Worksheet
    Dim signatureshp As Shape, chrt As ChartObject
    
    Sheets.Add(, Sheets(Sheets.Count)).Name = "Temp_signature"
    Set tempSheet = Sheets("Temp_signature")
    tempSheet.Select
    
    Set signatureshp = ThisWorkbook.Worksheets("mySignature").Shapes(SIGNATURE_SHAPE_NAME)
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

Private Sub GenerateReports(ByRef ws As Worksheet, ByVal clickedButtonName As String)
    Const REPORT_TEMPLATE As String = "Speaking Evaluation Template.docx"
    Const ERR_RESOURCES_FOLDER As String = "resourcesFolder"
    Const ERR_INCOMPLETE_RECORDS As String = "incompleteRecords"
    Const ERR_LOADING_WORD As String = "loadingWord"
    Const ERR_LOADING_TEMPLATE As String = "loadingTemplate"
    Const ERR_MISSING_SHAPES As String = "missingTemplateShapes"
    Const MSG_SAVE_FAILED As String = "exportFailed"
    Const MSG_ZIP_FAILED As String = "zipFailed"
    Const MSG_SUCCESS As String = "exportSuccessful"

    ' Objects to open Word and modify the template
    Dim wordApp As Object, wordDoc As Object
    
    ' Variables for determining the code path and important states
    Dim generateProcess As String, preexistingWordInstance As Boolean, saveResult As Boolean
    
    ' Strings for generating important messages for the user
    Dim resultMsg As String, msgToDisplay As String, msgTitle As String, msgType As Integer, msgResult As Variant, dialogSize As Integer
    
    ' Strings for tracking important filenames and filepaths
    Dim resourcesFolder As String, templatePath As String, savePath As String
    
    ' Numbers for iterating through student records and generate the reports
    Dim currentRow As Long, lastRow As Long, firstStudentRecord As Integer, i As Integer
    
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
            msgResult = DisplayMessage(msgToDisplay, vbExclamation, "Invalid Selection!")
        Exit Sub
    End Select
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Beginning report generation." & vbNewLine & _
                    "    Code Path: " & generateProcess
    #End If

    resourcesFolder = ThisWorkbook.Path & Application.PathSeparator & "Resources"
    ConvertOneDriveToLocalPath resourcesFolder
    
    #If Mac Then
        If Not RequestFileAndFolderAccess Then
            ' Create an error msg
            ' GoTo CleanUp
        End If
    #Else
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Checking for resources folder." & vbNewLine & _
                        "    Path: " & resourcesFolder
        #End If
        
        If Not DoesFolderExist(resourcesFolder) Then
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "    Folder not found. Attempting to create."
            #End If
            MkDir resourcesFolder
        End If
        
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Folder Created: " & DoesFolderExist(resourcesFolder)
        #End If
        
        If Not DoesFolderExist(resourcesFolder) Then
            resultMsg = ERR_RESOURCES_FOLDER
            GoTo CleanUp
        End If
    #End If

    If IsTemplateAlreadyOpen(resourcesFolder, REPORT_TEMPLATE, preexistingWordInstance) Then
        ' I can probably set an error msg and send this to CleanUp
        Exit Sub
    End If

    If Not VerifyRecordsAreComplete(ws, lastRow, firstStudentRecord) Then
        resultMsg = ERR_INCOMPLETE_RECORDS
        GoTo CleanUp
    End If

    templatePath = LocateTemplate(resourcesFolder, REPORT_TEMPLATE)
    If templatePath = "" Then
        ' Set an error msg
        GoTo CleanUp
    End If

    savePath = SetSaveLocation(ws, generateProcess)
    If savePath = "" Then
        ' Set an error msg
        GoTo CleanUp
    End If

    If Not LoadWord(wordApp, wordDoc, templatePath) Then
        resultMsg = ERR_LOADING_WORD
        GoTo CleanUp
    End If
    
    If wordDoc Is Nothing Then
        resultMsg = ERR_LOADING_TEMPLATE
        GoTo CleanUp
    End If
    
    If Not VerifyAllDocShapesExist(wordDoc) Then
        resultMsg = ERR_MISSING_SHAPES
        GoTo CleanUp
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Beginning to generate reports."
    #End If
    
    For currentRow = firstStudentRecord To lastRow
        #If PRINT_DEBUG_MESSAGES Then
            i = i + 1
            Debug.Print "    Current report: " & i & " of " & (lastRow - firstStudentRecord + 1)
        #End If
        ClearAllTextBoxes wordDoc
        WriteReport ws, wordApp, wordDoc, generateProcess, currentRow, savePath, saveResult
    Next currentRow
    
    If Not saveResult Then
        resultMsg = MSG_SAVE_FAILED
        GoTo CleanUp
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Save process complete."
    #End If
    
    KillWord wordApp, wordDoc, preexistingWordInstance
    resultMsg = MSG_SUCCESS
    
    If generateProcess = "FinalReports" Then
        ZipReports ws, savePath, saveResult
        If Not saveResult Then resultMsg = MSG_ZIP_FAILED
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
    
    If resultMsg <> "" Then msgResult = DisplayMessage(msgToDisplay, msgType, msgTitle, dialogSize)
    If Not wordApp Is Nothing Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Beginning final cleanup checks."
        #End If
        KillWord wordApp, wordDoc, preexistingWordInstance
    End If
End Sub

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
    useEmbeddedSignature = (Not ThisWorkbook.Sheets("mySignature").Shapes(SIGNATURE_SHAPE_NAME) Is Nothing)
    On Error GoTo 0
     
    If newImagePath = "" Then
        If useEmbeddedSignature Then
            ExportSignatureFromExcel SIGNATURE_SHAPE_NAME, newImagePath
        Else
            #If Mac Then
                newImagePath = AppleScriptTask(APPLE_SCRIPT_FILE, "FindSignature", signaturePath)
                If newImagePath = "" Then Exit Sub
                signatureFound = True
            #Else
                Select Case True
                    Case Dir(signaturePath & SIGNATURE_PNG_FILENAME) <> ""
                        newImagePath = signaturePath & SIGNATURE_PNG_FILENAME
                    Case Dir(signaturePath & SIGNATURE_JPG_FILENAME) <> ""
                        newImagePath = signaturePath & SIGNATURE_JPG_FILENAME
                    Case Else
                        Exit Sub
                End Select
            #End If
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

Private Sub KillWord(ByRef wordApp As Object, ByRef wordDoc As Object, ByVal preexistingWordInstance As Boolean)
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Attempting to close the open instance of MS Word."
    #End If
    
    On Error Resume Next
    If Not wordDoc Is Nothing Then
        wordDoc.Close SaveChanges:=False
        Set wordDoc = Nothing
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Attempting to close the template." & vbNewLine & _
                        "    Status: " & (wordDoc Is Nothing)
        #End If
    End If
    
    If preexistingWordInstance Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    A preexisting instance of MS Word was detected. For safety, it will not be closed."
        #End If
        Exit Sub
    End If
    
    If Not wordApp Is Nothing Then wordApp.Quit
    Set wordApp = Nothing
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Attempting to close MS Word." & vbNewLine & _
                    "    Status: " & (wordApp Is Nothing)
    #End If

    #If Mac Then
        Dim closeResult As String
        
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Attempting extra step required to complete close MS Word on MacOS."
        #End If
    
        closeResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CloseWord", closeResult)

        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Status: " & closeResult
        #End If
    #End If
    On Error GoTo 0
End Sub

Private Function LoadWord(ByRef wordApp As Object, ByRef wordDoc As Object, ByVal templatePath As String) As Boolean
    Dim openDoc As Object
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Attempting to open an instance of MS Word."
    #End If
    
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    Err.Clear
    On Error GoTo ErrorHandler
    
    ' Open a new instance of Word if needed
    #If Mac Then
        Dim appleScriptResult As String, msgToDisplay As String, msgResult As Variant
        
        If wordApp Is Nothing Then
            appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "LoadApplication", "Microsoft Word")
            
            #If PRINT_DEBUG_MESSAGES Then
                If appleScriptResult <> "" Then Debug.Print appleScriptResult
            #End If
            
            appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "IsAppLoaded", "Microsoft Word")
            
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "    " & appleScriptResult
            #End If
            
            Set wordApp = GetObject(, "Word.Application")
        End If
    #Else
        If wordApp Is Nothing Then Set wordApp = CreateObject("Word.Application")
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    MS Word loaded: " & (Not wordApp Is Nothing)
    #End If
    
    ' Make the process visible so users understand their computer isn't frozen
    wordApp.Visible = True
    wordApp.ScreenUpdating = True
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Visible: " & wordApp.Visible & vbNewLine & _
                    "    Show Updating: " & wordApp.ScreenUpdating
    #End If
    
    If Not wordApp Is Nothing Then
        Set wordDoc = wordApp.Documents.Open(templatePath)
        If Val(wordApp.Version) > 15 And wordDoc.AutoSaveOn Then
            On Error Resume Next
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "    Attempting to disable AutoSave."
            #End If
            wordDoc.AutoSaveOn = False
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "    AutoSave status: " & wordDoc.AutoSaveOn
            #End If
            On Error GoTo ErrorHandler
        End If
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Template loaded: " & (Not wordDoc Is Nothing)
    #End If
    
    LoadWord = (Not wordApp Is Nothing)
    Exit Function
ErrorHandler:
    #If Mac Then
        msgToDisplay = "An error occurred while trying to load Microsoft Word. This is usually a result of a quirk in MacOS. Try creating the reports again, and it should work fine." & vbNewLine & vbNewLine & _
                        "If the problem persists, please take a picture of the following error message and ask your team leader to send it to Warren at Bundang." & vbNewLine & vbNewLine & _
                        "VBA Error " & Err.Number & ": " & Err.Description
        #If Mac Then
            msgToDisplay = msgToDisplay & vbNewLine & "AppleScript Error: " & appleScriptResult
        #End If
        
        msgResult = DisplayMessage(msgToDisplay, vbOKOnly, "Error Loading Word", 470)
    #End If
    LoadWord = False
End Function

Private Function SaveToFile(ByRef wordDoc As Object, ByVal saveRoutine As String, ByVal savePath As String, ByVal fileName As String) As Boolean
    Dim tempFile As String, destFile As String
    Dim scriptResult As Boolean
    
    scriptResult = False
    
    On Error Resume Next
    If saveRoutine = "Proofs" Then
        tempFile = GetTempFilePath(fileName & ".docx")
        destFile = savePath & fileName & ".docx"
        
        ' Saving to the temp folder and then moving prevents a file permission issue with iCloud and having to wait for each upload with OneDrive
        #If Mac Then
            wordDoc.SaveAs2 fileName:=tempFile, FileFormat:=16, AddtoRecentFiles:=False, EmbedTrueTypeFonts:=True
            scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CopyFile", tempFile & APPLE_SCRIPT_SPLIT_KEY & destFile)
        #Else
            Dim fso As Object
            Set fso = CreateObject("Scripting.FileSystemObject")
            
            wordDoc.SaveAs2 fileName:=tempFile, FileFormat:=16, AddtoRecentFiles:=False, EmbedTrueTypeFonts:=True
            If fso.FileExists(tempFile) Then fso.CopyFile tempFile, destFile, True
        #End If
    Else
        #If Mac Then
            wordDoc.SaveAs2 fileName:=(savePath & fileName & ".pdf"), FileFormat:=17, AddtoRecentFiles:=False, EmbedTrueTypeFonts:=True
        #Else
            wordDoc.ExportAsFixedFormat OutputFileName:=(savePath & fileName & ".pdf"), ExportFormat:=17, BitmapMissingFonts:=True
        #End If
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        If Err.Number = 0 Then
            Debug.Print "        Report saved."
        Else
            Debug.Print "        Failed to save." & vbNewLine & _
                        "        Error Number: " & Err.Number & vbNewLine & _
                        "        Error Description: " & Err.Description
        End If
    #End If
    
    SaveToFile = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub WriteReport(ByRef ws As Object, ByRef wordApp As Object, ByRef wordDoc As Object, ByVal generateProcess As String, ByVal currentRow As Integer, ByVal savePath As String, ByRef saveResult As Boolean)
    Dim nativeTeacher As String, koreanTeacher As String, classLevel As String, classTime As String, evalDate As String
    Dim englishName As String, koreanName As String, grammarScore As String, pronunciationScore As String, fluencyScore As String
    Dim mannerScore As String, contentScore As String, effortScore As String, commentText As String, overallGrade As String
    Dim fileName As String
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "        Preparing report data."
    #End If
    
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
    
    ' fileName = koreanTeacher & "(" & classTime & ")" & " - " & koreanName & "(" & englishName & ")"
    fileName = koreanName & "(" & englishName & ")" & " - " & ws.Cells(4, 3).Value
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "        Report filename: " & fileName & vbNewLine & _
                    "        Saving to: " & savePath
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

Private Sub ZipReports(ByRef ws As Worksheet, ByVal savePath As String, ByRef saveResult As Boolean)
    Dim zipPath As Variant, zipName As Variant, pdfPath As Variant
    Dim errDescription As String
    
    On Error Resume Next
    If Right(savePath, 1) <> Application.PathSeparator Then savePath = savePath & Application.PathSeparator
    
    zipName = ws.Cells(3, 3).Value & " (" & ws.Cells(2, 3).Value & " " & ws.Cells(4, 3).Value & ").zip"
    zipPath = savePath & zipName
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Attempting to create a ZIP file of all generated reports for this class." & vbNewLine & _
                    "    Filename: " & zipName & vbNewLine & _
                    "    Path: " & savePath
    #End If
    
    #If Mac Then
        Dim scriptResult As String
        
        scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CreateZipFile", savePath & APPLE_SCRIPT_SPLIT_KEY & zipPath)
        
        If scriptResult <> "Success" Then
            errDescription = scriptResult
            saveResult = False
        Else
            saveResult = True
            scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "ClearPDFsAfterZipping", savePath)
        End If
    #Else
        Dim shellApp As Object
        
        ' Remove old copy if present
        If Len(Dir(zipPath)) > 0 Then Kill zipPath
        
        ' Create an empty ZIP file
        Open zipPath For Output As #1
        Print #1, "PK" & Chr(5) & Chr(6) & String(18, vbNullChar)
        Close #1
        
        Set shellApp = CreateObject("Shell.Application")
        pdfPath = Dir(savePath & "*.pdf") ' Only target PDF files
        
        Do While pdfPath <> ""
            shellApp.Namespace(zipPath).CopyHere savePath & pdfPath
            Application.Wait Now + TimeValue("0:00:01") ' Delay to allow compression
            ' Add a line to delete the PDF file
            pdfPath = Dir ' Get the next PDF file
        Loop
        
        Set shellApp = Nothing
        If Err.Number <> 0 Then errDescription = Err.Description
        saveResult = (Err.Number = 0)
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        If saveResult Then
            Debug.Print "    Zip file successfully created."
        Else
            Debug.Print "    There was an error creating the Zip file." & vbNewLine & _
                        "    Error: " & errDescription
        End If
    #End If
    On Error GoTo 0
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Student Records Validation
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function VerifyRecordsAreComplete(ByRef ws As Worksheet, ByRef lastRow As Long, ByRef firstStudentRecord As Integer) As Boolean
    Const CLASS_INFO_FIRST_ROW As Integer = 1
    Const CLASS_INFO_LAST_ROW As Integer = 6
    Const STUDENT_INFO_FIRST_ROW As Integer = 8
    Const STUDENT_INFO_FIRST_COL As Integer = 2
    Const STUDENT_INFO_LAST_COL As Integer = 10
    
    Dim currentRow As Long, currentColumn As Long
    Dim msgToDisplay As String, msgResult As Variant
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Verifying that student records are complete."
    #End If
    
    firstStudentRecord = STUDENT_INFO_FIRST_ROW ' Set here and passed back to keep things organized
    
    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.Count, STUDENT_INFO_FIRST_COL).End(xlUp).row
    On Error GoTo 0
    
    If lastRow < STUDENT_INFO_FIRST_ROW Then
        msgToDisplay = "No students were found!"
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    No students were found."
        #End If
        msgResult = DisplayMessage(msgToDisplay, vbOKOnly, "Error!", 160)
        VerifyRecordsAreComplete = False
        Exit Function
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Final student record entry: " & (lastRow - STUDENT_INFO_FIRST_ROW + 1) & vbNewLine & _
                    "    Beginning validation of entered records."
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
    Dim dataValue As String
    
    ' Static declarations to save a few cycles on subsequent runs if generating multiple classes
    Static validLevels As Variant, validDays As Variant, validTimes As Variant, gradeMapping As Variant, isDeclared As Boolean

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
                ' Map a numeric value to it's matching grade by its array index
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
    Dim currentRow As Integer, msgToDisplay As String, msgResult As Variant

    For currentRow = firstRow To lastRow
        If IsEmpty(ws.Cells(currentRow, 3).Value) Then
            msgToDisplay = "Class information incomplete." & vbNewLine & vbNewLine & _
                           "Missing: " & Left(ws.Cells(currentRow, 1).Value, Len(ws.Cells(currentRow, 1).Value) - 1)
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print msgToDisplay
            #End If
            msgResult = DisplayMessage(msgToDisplay, vbOKOnly, "Error!", 190)
            ValidateClassInfo = False
            Exit Function
        End If

        If currentRow >= 3 And currentRow <= 5 And Not ValidateData(ws.Cells(currentRow, 3), ws.Cells(currentRow, 1).Value) Then
            msgToDisplay = "Invalid value for " & Left(ws.Cells(currentRow, 1).Value, Len(ws.Cells(currentRow, 1).Value) - 1) & "." & vbNewLine & vbNewLine & _
                           "Would you like to ignore and continue?"
            If DisplayMessage(msgToDisplay, vbYesNo, "Error!", 250) = vbNo Then
                #If PRINT_DEBUG_MESSAGES Then
                    Debug.Print msgToDisplay
                #End If
                ValidateClassInfo = False
                Exit Function
            End If
        End If
    Next currentRow

    ValidateClassInfo = True
End Function

Private Function ValidateStudentInfo(ByRef ws As Worksheet, ByVal firstRow As Integer, ByVal lastRow As Integer, ByVal firstCol As Integer, ByVal lastCol As Integer) As Boolean
    Dim currentRow As Integer, currentColumn As Integer
    Dim msgToDisplay As String, msgResult As Variant, dialogSize As Integer
    
    For currentRow = firstRow To lastRow
        For currentColumn = firstCol To lastCol
            If IsEmpty(ws.Cells(currentRow, currentColumn).Value) Then
                msgToDisplay = "Student information incomplete." & vbNewLine & vbNewLine & _
                               "Missing data for student " & ws.Cells(currentRow, 1).Value & "'s "
                
                Select Case True
                    Case currentColumn = 2
                        msgToDisplay = msgToDisplay & ws.Cells(7, currentColumn).Value & "."
                    Case (currentColumn >= 4 And currentColumn <= 9)
                        msgToDisplay = msgToDisplay & "(" & ws.Cells(currentRow, 2).Value & ") " & _
                                       UCase(ws.Cells(7, currentColumn).Value) & " score."
                    Case currentColumn = 10
                        msgToDisplay = msgToDisplay & "(" & ws.Cells(currentRow, 2).Value & ") COMMENT."
                End Select

                dialogSize = 250
                GoTo ErrorHandler
            End If
            
            If currentColumn >= 4 And Not ValidateData(ws.Cells(currentRow, currentColumn), ws.Cells(7, currentColumn).Value) Then
                If currentColumn <> 10 Then
                    msgToDisplay = "Invalid value entered for student " & ws.Cells(currentRow, 1).Value & "'s (" & ws.Cells(currentRow, 2).Value & ") " & UCase(ws.Cells(7, currentColumn).Value) & " score."
                    dialogSize = 300
                Else
                    msgToDisplay = "The COMMENT for student " & ws.Cells(currentRow, 1).Value & "'s (" & ws.Cells(currentRow, 2).Value & ") is too long. Please try to shorten it by " & _
                                    Len(ws.Cells(currentRow, currentColumn).Value) - 315 & " or more characters."
                    dialogSize = 330
                End If
                GoTo ErrorHandler
            End If
        Next currentColumn
    Next currentRow
    
    ValidateStudentInfo = True
    Exit Function
    
ErrorHandler:
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print msgToDisplay
    #End If
    msgResult = DisplayMessage(msgToDisplay, vbExclamation, "Error!", dialogSize)
    ValidateStudentInfo = False
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Report Template Management
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub ClearAllTextBoxes(wordDoc As Object)
    Dim shp As Object, grpItem As Object
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "        Clearing text from textboxes."
    #End If
    
    For Each shp In wordDoc.Shapes
        If shp.Type = msoGroup Then
            For Each grpItem In shp.GroupItems
                If grpItem.Type = msoTextBox Or grpItem.Type = msoAutoShape Then
                    grpItem.TextFrame.TextRange.Text = ""
                End If
            Next grpItem
        End If
    Next shp
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "        Complete."
    #End If
End Sub

Private Function DownloadReportTemplate(ByVal templatePath As String) As Boolean
    Const REPORT_TEMPLATE_URL As String = "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/Speaking%20Evaluation%20Template.docx"
    Dim downloadResult As Boolean
    
    #If Mac Then
        On Error Resume Next
        downloadResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DownloadFile", templatePath & APPLE_SCRIPT_SPLIT_KEY & REPORT_TEMPLATE_URL)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print IIf(Err.Number = 0, "    Download successful.", "    Error: " & Err.Description)
        #End If
        
        If downloadResult Then downloadResult = RequestFileAndFolderAccess(templatePath)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    File access " & IIf(downloadResult, "granted.", "denied.")
        #End If
        On Error GoTo 0
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

Private Function LocateTemplate(ByVal resourcesFolder As String, ByVal REPORT_TEMPLATE As String) As String
    Dim templatePath As String, tempTemplatePath As String
    Dim msgToDisplay As String, msgResult As Variant
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Attempting to load the Speaking Evaluation Template.docx."
    #End If
    
    templatePath = resourcesFolder & Application.PathSeparator & REPORT_TEMPLATE
    tempTemplatePath = GetTempFilePath(REPORT_TEMPLATE)
    
    DeleteFile tempTemplatePath ' Removing existing file to avoid problems overwriting

    If Not VerifyTemplateHash(templatePath) And Not DownloadReportTemplate(templatePath) Then
        msgToDisplay = "No template was found. Process canceled."
        msgResult = DisplayMessage(msgToDisplay, vbOKOnly, "Template Not Found", 150)
        LocateTemplate = ""
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Unable to locate a copy of the template."
        #End If
        Exit Function
    End If

    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Valid template found."
    #End If
    
    If MoveFile(templatePath, tempTemplatePath) Then
        LocateTemplate = tempTemplatePath
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Loading temporary copy."
        #End If
    Else
        LocateTemplate = templatePath
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Failed to make a temporary copy. Using resources copy directly."
        #End If
    End If
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
                       
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Verifying all template shapes are present."
    #End If
    
    For i = LBound(shapeNames) To UBound(shapeNames)
        If Not WordDocShapeExists(wordDoc, shapeNames(i)) Then
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "    Missing shape: " & shapeNames(i)
            #End If
            
            msgToDisplay = "There is a critical error with the template. Please redownload a copy of the original and try again."
            msgResult = DisplayMessage(msgToDisplay, vbOKOnly, "Error!", 300)
            VerifyAllDocShapesExist = False
            Exit Function
        End If
    Next i
                       
    VerifyAllDocShapesExist = True
End Function

Private Function VerifyTemplateHash(ByVal templatePath As String) As Boolean
    Const TEMPLATE_HASH As String = "C0343895A881DF739B2B974635A100A6"
    
    #If Mac Then
        VerifyTemplateHash = AppleScriptTask(APPLE_SCRIPT_FILE, "CompareMD5Hashes", templatePath & APPLE_SCRIPT_SPLIT_KEY & TEMPLATE_HASH)
    #Else
        Dim objShell As Object, shellOutput As String
        If Dir(templatePath) <> "" Then
            On Error GoTo ErrorHandler
            Set objShell = CreateObject("WScript.Shell")
            shellOutput = objShell.Exec("cmd /c certutil -hashfile """ & templatePath & """ MD5").StdOut.ReadAll
            VerifyTemplateHash = (LCase(TEMPLATE_HASH) = LCase(Trim(Split(shellOutput, vbCrLf)(1))))
        Else
            VerifyTemplateHash = False
        End If
    #End If
CleanUp:
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
    Resume CleanUp
End Function

Private Function WordDocShapeExists(ByRef wordDoc As Object, ByVal shapeName As String) As Boolean
    Dim shp As Object, grpItem As Object
    
    On Error Resume Next
    For Each shp In wordDoc.Shapes
        If shp.Type = msoGroup Then
            For Each grpItem In shp.GroupItems
                If grpItem.Name = shapeName Then
                    WordDocShapeExists = True
                    
                    #If PRINT_DEBUG_MESSAGES Then
                        Debug.Print "    " & shapeName & ": Present"
                    #End If
                    
                    Exit Function
                End If
            Next grpItem
        End If
    Next shp
    On Error GoTo 0
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    " & shapeName & ": Missing"
    #End If
    
    WordDocShapeExists = False
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Configuration and File Management Routines
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Windows and MacOS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub ConvertOneDriveToLocalPath(ByRef selectedPath As Variant)
    Dim i As Integer
    
    ' Cloud storage apps like OneDrive sometimes complicate where/how files are saved. Below is a reference
    ' to track and help add support for additionalcloud storage providers.
    
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

Private Sub CreateSaveFolder(ByRef filePath As String)
    If Right(filePath, 1) = Application.PathSeparator Then
        filePath = Left(filePath, Len(filePath) - 1)
    End If

    On Error Resume Next
    #If Mac Then
        Dim scriptResult As Boolean
        scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "CreateFolder", filePath)
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

Private Sub DeleteFile(ByVal filePath As String)
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Deleting template from temporary files folder."
    #End If
    
    #If Mac Then
        Dim appleScriptResult As Boolean
        
        appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", filePath)
        If appleScriptResult Then appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", filePath)
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

        scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "ClearFolder", filePath)
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

Public Function DisplayMessage(ByVal messageText As String, ByVal messageType As Integer, ByVal messageTitle As String, Optional ByVal dialogWidth As Integer = 250) As Variant
    #If Mac Then
        Dim dialogType As String, iconType As String
        Dim dialogParameters As String, dialogResult As Variant, messageDisplayed As Boolean
        Dim lastError As String
        Dim i As Integer
        
        ' Button types for bitwise comparison
        Const BUTTON_OK_ONLY As Integer = 0
        Const BUTTON_OK_CANCEL As Integer = 1
        Const BUTTON_RETRY_CANCEL As Integer = 2
        Const BUTTON_YES_NO As Integer = 4
        Const BUTTON_YES_NO_CANCEL As Integer = 8
        
        ' Icon types for bitwise comparison
        Const ICON_CRITICAL As Integer = 16
        Const ICON_QUESTION As Integer = 32
        Const ICON_EXCLAMATION As Integer = 48
        Const ICON_INFORMATION As Integer = 64

        If AreEnhancedDialogsEnabled Then
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "Attempting to display message via Dialog Toolkit Plus." & vbNewLine & _
                            "    Message: " & messageText
            #End If
            
            ' Determine buttons to display
            Select Case True
                Case (messageType And BUTTON_OK_ONLY) = BUTTON_OK_ONLY
                    dialogType = "OkOnly"
                Case (messageType And BUTTON_OK_CANCEL) = BUTTON_OK_CANCEL
                    dialogType = "OkCancel"
                Case (messageType And BUTTON_RETRY_CANCEL) = BUTTON_RETRY_CANCEL
                    dialogType = "RetryCancel"
                Case (messageType And BUTTON_YES_NO) = BUTTON_YES_NO
                    dialogType = "YesNo"
                Case (messageType And BUTTON_YES_NO_CANCEL) = BUTTON_YES_NO_CANCEL
                    dialogType = "YesNoCancel"
                Case Else
                    dialogType = "OkOnly"
            End Select
            
            ' Determine icon to display
            Select Case True
                Case (messageType And ICON_CRITICAL) = ICON_CRITICAL
                    iconType = "CriticalIcon"
                Case (messageType And ICON_QUESTION) = ICON_QUESTION
                    iconType = "QuestionIcon"
                Case (messageType And ICON_EXCLAMATION) = ICON_EXCLAMATION
                    iconType = "ExclamationIcon"
                Case (messageType And ICON_INFORMATION) = ICON_INFORMATION
                    iconType = "InformationIcon"
                Case Else
                    iconType = "OtherIcon"
            End Select
                    
            ' Update SpeakingEvals.scpt to supoport:
            ' dialogParameters = messageText & APPLE_SCRIPT_SPLIT_KEY & dialogType & APPLE_SCRIPT_SPLIT_KEY & iconType & APPLE_SCRIPT_SPLIT_KEY & messageTitle & APPLE_SCRIPT_SPLIT_KEY & dialogWidth
            dialogParameters = messageText & APPLE_SCRIPT_SPLIT_KEY & dialogType & APPLE_SCRIPT_SPLIT_KEY & messageTitle & APPLE_SCRIPT_SPLIT_KEY & dialogWidth
            
            On Error Resume Next
            Do While Not messageDisplayed
                dialogResult = AppleScriptTask("DialogDisplay.scpt", "DisplayDialog", dialogParameters)
                messageDisplayed = (dialogResult <> "")
                i = i + 1
                
                #If PRINT_DEBUG_MESSAGES Then
                    If Err.Number <> 0 Then lastError = Err.Number & " - " & Err.Description
                #End If
                
                If i >= 10 Then
                    dialogResult = MsgBox(messageText, messageType, messageTitle)
                    messageDisplayed = True
                End If
            Loop
            On Error GoTo 0
            
            #If PRINT_DEBUG_MESSAGES Then
                If lastError = "" Then lastError = "N/A"
                Debug.Print "    Number of attempts: " & i & vbNewLine & _
                            "    Final error: " & lastError
            #End If
            
            DisplayMessage = dialogResult
            Exit Function
        End If
    #End If
    
    DisplayMessage = MsgBox(messageText, messageType, messageTitle)
End Function

Private Function DoesFolderExist(ByVal filePath As String) As Boolean
    #If Mac Then
        DoesFolderExist = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFolderExist", filePath)
    #Else
        DoesFolderExist = (Dir(filePath, vbDirectory) <> "")
    #End If
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

Private Function GetTempFilePath(ByVal fileName As String) As String
    #If Mac Then
        GetTempFilePath = Environ("TMPDIR") & fileName
    #Else
        GetTempFilePath = Environ("TEMP") & Application.PathSeparator & fileName
    #End If
End Function

Private Function IsTemplateAlreadyOpen(ByVal resourcesFolder As String, ByVal REPORT_TEMPLATE As String, ByRef preexistingWordInstance As Boolean) As Boolean
    Dim wordApp As Object, wordDoc As Object
    Dim templatePath As String, templateIsOpen As Boolean
    Dim pathOfOpenDoc As String
    Dim msgToDisplay As String
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Checking if a copy of the Speaking Evaluations Report template is already open."
    #End If
    
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    Err.Clear
    
    If Not wordApp Is Nothing Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Found an open instance of MS Word." & vbNewLine & _
                        "    Checking if template is open."
        #End If
        
        preexistingWordInstance = True
        templatePath = resourcesFolder & Application.PathSeparator & REPORT_TEMPLATE
        
        For Each wordDoc In wordApp.Documents
            pathOfOpenDoc = wordDoc.FullName
            ConvertOneDriveToLocalPath pathOfOpenDoc
            If StrComp(pathOfOpenDoc, templatePath, vbTextCompare) = 0 Then
                templateIsOpen = True
                #If PRINT_DEBUG_MESSAGES Then
                    Debug.Print "    Open instance of the template found. Asking if user wishes to automatically close and continue."
                #End If
                 msgToDisplay = "An open instance of MS Word has been detected. Please save any open files before continuing." & vbNewLine & vbNewLine & _
                                "Click OK to automatically close Word and continue, or click Cancel to finish and save your work."
                If DisplayMessage(msgToDisplay, vbOKCancel + vbCritical, "Error Loading Word", 310) = vbOK Then
                    wordDoc.Close SaveChanges:=False
                    templateIsOpen = False
                    #If PRINT_DEBUG_MESSAGES Then
                        Debug.Print "    Open instance has been closed."
                    #End If
                End If
            End If
        Next wordDoc
    End If
    On Error GoTo 0
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Open instance: " & templateIsOpen
    #End If
    
    Set wordDoc = Nothing
    Set wordApp = Nothing
    IsTemplateAlreadyOpen = templateIsOpen
End Function

Private Function MoveFile(ByVal initialPath As String, ByVal destinationPath As String) As Boolean
    Dim moveSuccessful As Boolean
    
    On Error Resume Next
    #If Mac Then
        moveSuccessful = AppleScriptTask(APPLE_SCRIPT_FILE, "CopyFile", initialPath & APPLE_SCRIPT_SPLIT_KEY & destinationPath)
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
        End If
    #End If
    
    Err.Clear
    On Error GoTo 0
    MoveFile = moveSuccessful
End Function

Private Function SetSaveLocation(ByRef ws As Object, ByVal saveRoutine As String) As String
    Dim filePath As String
    
    filePath = ThisWorkbook.Path & Application.PathSeparator & GenerateSaveFolderName(ws) & Application.PathSeparator
    ConvertOneDriveToLocalPath filePath
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Setting save location for generated reports." & vbNewLine & _
                    "    Save location: " & filePath
    #End If

    If DoesFolderExist(filePath) Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Path already exists. Clearing out old files."
        #End If
        DeleteExistingFolder filePath
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Creating save path."
    #End If
    
    CreateSaveFolder filePath
    #If Mac Then
        Dim permissionGranted As Boolean
        permissionGranted = RequestFileAndFolderAccess(filePath)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print IIf(permissionGranted, "    Folder access granted. Continuing with process", "    Folder access denied. Cannot continue.")
        #End If
        If Not permissionGranted Then
            ' Add a savePath permission denied value
            SetSaveLocation = ""
            Exit Function
        End If
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Saving reports in: " & vbNewLine & _
                    "    " & filePath
    #End If
    
    SetSaveLocation = filePath
End Function

#If Mac Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MacOS Only
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function AreAppleScriptsInstalled(Optional ByVal recheckStatus As Boolean = False) As Boolean
    Dim libraryScriptsFolder As String, resourcesFolder As String, isAppleScriptInstalled As Boolean
    Dim isDialogToolkitInstalled As Boolean, statusHasBeenChecked As Boolean, scriptResult As Boolean
    
    isAppleScriptInstalled = CheckForAppleScript()
    
    If isAppleScriptInstalled Then
        If Not recheckStatus Then CheckForAppleScriptUpdate
        
        libraryScriptsFolder = "/Users/" & Environ("USER") & "/Library/Script Libraries"
        resourcesFolder = ThisWorkbook.Path & "/Resources"
        ConvertOneDriveToLocalPath resourcesFolder

        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Locating Dialog Toolkit Plus.scptd" & vbNewLine & _
                        "    Searching: " & libraryScriptsFolder
        #End If

        If Not recheckStatus Then
            ' When first opened, only check for Dialog Toolkit Plus if the folder has been previously created
            scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFolderExist", libraryScriptsFolder)
            If scriptResult Then isDialogToolkitInstalled = CheckForDialogToolkit(resourcesFolder)
        Else
            isDialogToolkitInstalled = CheckForDialogToolkit(resourcesFolder)
        End If

        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Installed: " & isDialogToolkitInstalled
        #End If

        If isDialogToolkitInstalled Then
            isDialogToolkitInstalled = CheckForDialogDisplayScript(resourcesFolder)
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "Attempting to install DialogDisplay.scpt" & vbNewLine & _
                            "    Installed: " & isDialogToolkitInstalled
            #End If
        End If
    Else
        isDialogToolkitInstalled = False
    End If

    SetVisibilityOfMacSettingsShapes isAppleScriptInstalled, isDialogToolkitInstalled

    AreAppleScriptsInstalled = isAppleScriptInstalled
End Function

Private Function AreEnhancedDialogsEnabled() As Boolean
    AreEnhancedDialogsEnabled = ThisWorkbook.Sheets("MacOS Users").Shapes("Button_EnhancedDialogs_Enable").Visible
End Function

Private Function CheckForAppleScript() As Boolean
    Dim appleScriptPath As String, appleScriptStatus As Boolean
    
    appleScriptPath = "/Users/" & Environ("USER") & "/Library/Application Scripts/com.microsoft.Excel/" & APPLE_SCRIPT_FILE
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Locating " & APPLE_SCRIPT_FILE & vbNewLine & _
                    "    Searching: " & appleScriptPath
    #End If
    
    On Error Resume Next
    appleScriptStatus = (Dir(appleScriptPath, vbDirectory) = APPLE_SCRIPT_FILE)
    On Error GoTo 0
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Found: " & appleScriptStatus
    #End If
    
    CheckForAppleScript = appleScriptStatus
End Function

Private Sub CheckForAppleScriptUpdate()
    Dim scriptFolder As String, destinationPath As String
    Dim currentScriptVersion As Long, downloadedScriptVersion As Long
    Dim appleScriptResult As Boolean
    
    Const APPLE_SCRIPT_URL As String = "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/SpeakingEvals.scpt"
    Const OLD_APPLE_SCRIPT As String = "SpeakingEvals-Old.scpt"
    Const TMP_APPLE_SCRIPT As String = "SpeakingEvals-Tmp.scpt"
    
    scriptFolder = "/Users/" & Environ("USER") & "/Library/Application Scripts/com.microsoft.Excel/"
    destinationPath = scriptFolder & TMP_APPLE_SCRIPT
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Checking if an update is available for SpeakingEvals.scpt."
    #End If
    
    On Error GoTo ErrorHandler
    
    appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DownloadFile", destinationPath & APPLE_SCRIPT_SPLIT_KEY & APPLE_SCRIPT_URL)
    If Not appleScriptResult Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Unable to download new " & APPLE_SCRIPT_FILE
        #End If
        GoTo ErrorHandler
    End If
    
    currentScriptVersion = AppleScriptTask(APPLE_SCRIPT_FILE, "GetScriptVersionNumber", "")
    downloadedScriptVersion = AppleScriptTask(TMP_APPLE_SCRIPT, "GetScriptVersionNumber", "")
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Installed Version: " & currentScriptVersion & vbNewLine & _
                    "    Online Version:    " & downloadedScriptVersion
    #End If
    
    If downloadedScriptVersion <= currentScriptVersion Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Installed version is up-to-date."
        #End If
        GoTo CleanUp
    End If
    
    appleScriptResult = AppleScriptTask(TMP_APPLE_SCRIPT, "RenameFile", scriptFolder & APPLE_SCRIPT_FILE & APPLE_SCRIPT_SPLIT_KEY & scriptFolder & OLD_APPLE_SCRIPT)
    If appleScriptResult Then appleScriptResult = AppleScriptTask(OLD_APPLE_SCRIPT, "RenameFile", scriptFolder & TMP_APPLE_SCRIPT & APPLE_SCRIPT_SPLIT_KEY & scriptFolder & APPLE_SCRIPT_FILE)
    If appleScriptResult Then appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", scriptFolder & OLD_APPLE_SCRIPT)
    If Not appleScriptResult Then GoTo ErrorHandler
    
    #If PRINT_DEBUG_MESSAGES Then
        If appleScriptResult Then Debug.Print "    Update complete."
    #End If
    
CleanUp:
    #If PRINT_DEBUG_MESSAGES Then
        If appleScriptResult Then Debug.Print "    Beginning clean up process."
    #End If
    
    On Error Resume Next
    appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", scriptFolder & TMP_APPLE_SCRIPT)
    If appleScriptResult Then
        appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", scriptFolder & TMP_APPLE_SCRIPT)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Removing temporary update file: " & IIf(appleScriptResult, "Successful", "Failed")
        #End If
    End If
    
    appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", scriptFolder & OLD_APPLE_SCRIPT)
    If appleScriptResult Then
        appleScriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DeleteFile", scriptFolder & OLD_APPLE_SCRIPT)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Removing old version: " & IIf(appleScriptResult, "Successful", "Failed")
        #End If
    End If
    On Error GoTo 0
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Finished clean up."
    #End If
    Exit Sub
    
ErrorHandler:
    #If PRINT_DEBUG_MESSAGES Then
        If Err.Number <> 0 Then Debug.Print "Error during the update process."
        If Err.Description <> "" Then Debug.Print "Error: " & Err.Description
    #End If
    Resume CleanUp
End Sub

Private Function CheckForDialogToolkit(ByVal resourcesFolder As String) As Boolean
    Dim scriptResult As Boolean, libraryScriptsPath As String
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Checking for presence of Dialog Toolkit Plus." & vbNewLine & _
                    "    Local resources: " & resourcesFolder
    #End If
    
    libraryScriptsPath = AppleScriptTask(APPLE_SCRIPT_FILE, "CheckForScriptLibrariesFolder", "paramString")
    If libraryScriptsPath <> "" Then scriptResult = RequestFileAndFolderAccess(libraryScriptsPath)
    If scriptResult Then scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "InstallDialogToolkitPlus", resourcesFolder)
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Toolkit Status: " & scriptResult
    #End If
    
    CheckForDialogToolkit = scriptResult
End Function

Private Function CheckForDialogDisplayScript(ByVal resourcesFolder As String) As Boolean
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Checking for presence of DialogDisplay.scpt."
    #End If
        
    CheckForDialogDisplayScript = AppleScriptTask(APPLE_SCRIPT_FILE, "InstallDialogDisplayScript", resourcesFolder)
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Status: " & CheckForDialogDisplayScript
    #End If
End Function

Private Sub RemoveDialogToolKit(ByVal resourcesFolder As String)
    Dim scriptResult As Boolean
        
    If CheckForAppleScript() Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Removing Dialog ToolKit Plus from ~/Library/Script Libraries" & vbNewLine & _
                        "    A local copy will be stored in: " & resourcesFolder
        #End If
            
        scriptResult = AppleScriptTask(APPLE_SCRIPT_FILE, "UninstallDialogToolkitPlus", resourcesFolder)
            
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Result: " & scriptResult
        #End If
    End If
End Sub

Private Function RequestFileAndFolderAccess(Optional ByVal filePath As Variant = "") As Boolean
    Dim workingFolder As Variant, resourcesFolder As Variant, excelTempFolder As Variant, wordTempFolder As Variant
    Dim filePermissionCandidates As Variant, pathToRequest As Variant
    Dim fileAccessGranted As Boolean, allAccessHasBeenGranted As Boolean
    Dim i As Integer

    Select Case filePath
        Case ""
            workingFolder = ThisWorkbook.Path
            ConvertOneDriveToLocalPath workingFolder
            resourcesFolder = workingFolder & "/Resources"
            excelTempFolder = Environ("TMPDIR")
            wordTempFolder = Replace(excelTempFolder, "Excel", "Word")
            filePermissionCandidates = Array(workingFolder, resourcesFolder, excelTempFolder, wordTempFolder)
        Case Else
            ConvertOneDriveToLocalPath filePath ' Seems to be not needed?
            filePermissionCandidates = Array(filePath)
    End Select

    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Requesting access to: "
    #End If

    For i = LBound(filePermissionCandidates) To UBound(filePermissionCandidates)
        pathToRequest = Array(filePermissionCandidates(i))
        fileAccessGranted = GrantAccessToMultipleFiles(pathToRequest)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    " & filePermissionCandidates(i) & vbNewLine & _
                        "    Access granted: " & fileAccessGranted
        #End If
        allAccessHasBeenGranted = fileAccessGranted
        If Not fileAccessGranted Then Exit For
    Next i

    RequestFileAndFolderAccess = allAccessHasBeenGranted
End Function

Private Sub SetVisibilityOfMacSettingsShapes(ByVal isAppleScriptInstalled As Boolean, ByVal isDialogToolkitInstalled As Boolean)
    Dim ws As Worksheet, shps As Shapes
    Dim enhancedDialogsStatus As String
    Dim enhancedDialogsAreDisabled As Boolean
    
    Set ws = ThisWorkbook.Sheets("MacOS Users")
    Set shps = ws.Shapes
    
    shps("Button_SpeakingEvalsScpt_Missing").Visible = Not isAppleScriptInstalled
    shps("Button_SpeakingEvalsScpt_Installed").Visible = isAppleScriptInstalled
    shps("Button_DialogToolkit_Missing").Visible = Not isDialogToolkitInstalled
    shps("Button_DialogToolkit_Installed").Visible = isDialogToolkitInstalled
    
    enhancedDialogsStatus = ws.Cells(1, 1).Value
    enhancedDialogsAreDisabled = Not isDialogToolkitInstalled Or enhancedDialogsStatus = "Enhanced Dialogs: Disabled" Or enhancedDialogsStatus = ""
    
    shps("Button_EnhancedDialogs_Disable").Visible = enhancedDialogsAreDisabled
    shps("Button_EnhancedDialogs_Enable").Visible = Not enhancedDialogsAreDisabled
    
    If enhancedDialogsAreDisabled And enhancedDialogsStatus <> "Enhanced Dialogs: Disabled" Then
        ws.Unprotect
        ws.Cells(1, 1).Value = "Enhanced Dialogs: Disabled"
        ws.Protect
        ws.EnableSelection = xlUnlockedCells
    End If
End Sub

Private Sub RemindUserToInstallSpeakingEvalsScpt()
    Const msgToDisplay As String = "SpeakingEvals.scpt must be installed in order to generate reports. Please run the terminal command on the ""MacOs Users"" sheet to install it and try again."
    Dim msgResult As Integer

    msgResult = DisplayMessage(msgToDisplay, vbExclamation, "Invalid Selection!")
    ThisWorkbook.Sheets("MacOS Users").Activate
End Sub

Private Sub ToogleMacSettingsButtons(ByRef ws As Worksheet, ByVal clickedButtonName As String)
    Const SCRIPT_ENABLED As String = "Enhanced Dialogs: Enabled"
    Const SCRIPT_DISABLED As String = "Enhanced Dialogs: Disabled"
    
    Dim shps As Shapes
    Dim installedStatus As Boolean

    Set shps = ws.Shapes

    If shps("Button_DialogToolkit_Missing").Visible Then
        installedStatus = AreAppleScriptsInstalled(True)
        
        ' Button_EnhancedDialogs_Enable isn't visible yet (a quirk of the safety checks, but expected behaviour), so
        ' we need to check the visibility of Button_DialogToolkit_Installed to determine installation success.
        If Not shps("Button_DialogToolkit_Installed").Visible Then
            shps("Button_EnhancedDialogs_Disable").Visible = False
            shps("Button_EnhancedDialogs_Enable").Visible = True
            Exit Sub
        End If
    End If

    Select Case clickedButtonName
        Case "Button_EnhancedDialogs_Enable"
            shps("Button_EnhancedDialogs_Disable").Visible = True
            shps("Button_EnhancedDialogs_Enable").Visible = False
        Case "Button_EnhancedDialogs_Disable"
            shps("Button_EnhancedDialogs_Enable").Visible = True
            shps("Button_EnhancedDialogs_Disable").Visible = False
    End Select
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Updating persistant status value."
    #End If
    
    ws.Unprotect
    ws.Cells(1, 1).Value = IIf(ws.Shapes("Button_EnhancedDialogs_Enable").Visible, SCRIPT_ENABLED, SCRIPT_DISABLED)
    ws.Protect
    ws.EnableSelection = xlUnlockedCells
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Value: """ & ws.Cells(1, 1).Value & """"
    #End If
End Sub
#Else
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Windows Only
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
        Debug.Print IIf(checkResult, "    Installed.", "    Not installed. Falling back to .Net.")
    #End If
    
    CheckForCurl = checkResult
CleanUp:
    If Not objExec Is Nothing Then Set objExec = Nothing
    If Not objShell Is Nothing Then Set objShell = Nothing
    Exit Function
ErrorHandler:
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Error while checking for curl.exe: " & Err.Description
    #End If
    CheckForCurl = False
    Resume CleanUp
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
        If Not DownloadUsingCurl Then Debug.Print "    curl download failed for " & REPORT_TEMPLATE_URL
    #End If
    
    Set objShell = Nothing
    Set fso = Nothing
    On Error GoTo 0
End Function

Private Function DownloadUsingDotNet35(ByVal templatePath As String, ByVal REPORT_TEMPLATE_URL As String) As Boolean
    Dim xmlHTTP As Object, fileStream As Object
    
    On Error Resume Next
    Set xmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Set fileStream = CreateObject("ADODB.Stream")
    
    xmlHTTP.Open "Get", REPORT_TEMPLATE_URL, False
    xmlHTTP.Send
    
    If xmlHTTP.status = 200 Then
        fileStream.Open
        fileStream.Type = 1 ' Binary
        fileStream.Write xmlHTTP.responseBody
        fileStream.SaveToFile templatePath, 2 ' Overwrite existing, if somehow present
        fileStream.Close
        DownloadUsingDotNet35 = True
    Else
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "HTTP request failed. Status: " & xmlHTTP.status & " - " & xmlHTTP.StatusText
        #End If
        DownloadUsingDotNet35 = False
    End If
    
    Set xmlHTTP = Nothing
    Set fileStream = Nothing
    On Error GoTo 0
End Function
#End If
