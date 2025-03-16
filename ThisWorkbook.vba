''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Global Declarations and Main Sub
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
#Const PRINT_DEBUG_MESSAGES = True
Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
Const APPLE_SCRIPT_SPLIT_KEY = "-,-"

Sub Main()
    Dim ws As Worksheet, clickedButtonName As String
    
    Set ws = ActiveSheet
    
    With Application
        clickedButtonName = .Caller
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Beginning tasks." & vbNewLine & _
                    "    Active Worksheet = " & ws.Name & vbNewLine & _
                    "    Button Pressed: " & clickedButtonName
    #End If
    
    ' Check system requirements
    #If Mac Then
        Dim msgToDiplay As String
        
        If Not AreAppleScriptsInstalled(True) Then
            RemindUserToInstallSpeakingEvalsScpt
            Exit Sub
        End If
    #Else
        'Dim requirementsMet As Boolean, rebootRequired As Boolean
        'Dim msgResult As Integer
        
        'Const REBOOT_MSG As String = "Please reboot your computer and try again."
        'Const REBOOT_MSG_MSGTYPE As Integer = vbOKOnly + vbExclamation
        'Const REBOOT_MSG_TITLE As String = "Reboot Required!"
        
        'Const NO_ARCHIVE_TOOL_MSG As String = ""
        'Const NO_ARCHIVE_TOOL_MSGTYPE As Integer = vbYesNo + vbExclamation
        'Const NO_ARCHIVE_TOOL_TITLE As String = "Reboot Required!"
        
        'requirementsMet = AreKoreanFilenamesSupported(rebootRequired)
        
        'If rebootRequired Then
        '    msgResult = DisplayMessage(REBOOT_MSG, REBOOT_MSG_MSGTYPE, REBOOT_MSG_TITLE)
        '    Exit Sub
        'End If
        
        'If Not requirementsMet Then
        '    If FindPathToArchiveTool = "" Then
        '        msgResult = DisplayMessage(NO_ARCHIVE_TOOL_MSG, NO_ARCHIVE_TOOL_MSGTYPE, NO_ARCHIVE_TOOL_TITLE)
        '        If msgResult = vbYes Then
        '            ' Decide / Test what to do in this case
        '            ' Let the user simple make the PDFs or abort?
        '        End If
        '    End If
        'End If
    #End If
    
    On Error GoTo ReenableEvents
    
    Select Case clickedButtonName
        #If Mac Then
        Case "Button_EnhancedDialogs_Enable", "Button_EnhancedDialogs_Disable"
            ToogleMacSettingsButtons ws, clickedButtonName
        #End If
        Case "Button_GenerateReports", "Button_GenerateProofs"
            GenerateReports ws, clickedButtonName
            ws.Activate ' Ensure the right worksheet is being shown when finished.
        Case "Button_RepairLayout"
            ws.Unprotect
            SetLayoutClassRecords ThisWorkbook, ws
            ws.Protect
    End Select
    
ReenableEvents:
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  Worksheet & Data Validation
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Workbook_Open()
    Const CURL_COMMAND_TEXT As String = "curl -L -o ~/Library/Application\ Scripts/com.microsoft.Excel/SpeakingEvals.scpt https://github.com/papercutter0324/SpeakingEvals/raw/main/SpeakingEvals.scpt"
    Dim wb As Workbook, ws As Worksheet, shps As Shapes
    Dim startupMessageToDisplay As String
    Dim startTime As Double
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Beginning start-up self-checks."
    #End If
    
    Set wb = ThisWorkbook
    wb.Sheets("Instructions").Activate
    
    #If Mac Then
        wb.Sheets("MacOS Users").Visible = xlSheetVisible
    #Else
        wb.Sheets("MacOS Users").Visible = xlSheetHidden
    #End If
    
    startTime = Timer
    While Timer - startTime < 2
        DoEvents
    Wend
        
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "    Application.EnableEvents = " & Application.EnableEvents & _
                    "    Application.ScreenUpdating = " & Application.ScreenUpdating
    #End If
    
    startupMessageToDisplay = "Initial"
    DisplayStartupMessage startupMessageToDisplay
    
    VerifySheetNames wb
    
    On Error GoTo ReenableEvents
    For Each ws In wb.Worksheets
        With ws
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "Validating Layout: " & .Name
            #End If
            
            .Unprotect
            
            Select Case .Name
                Case "Instructions"
                    .Cells(1, 3).Select
                    SetLayoutInstructions wb, ws
                Case "MacOS Users"
                    Set shps = .Shapes
                    shps("cURL_Command").TextFrame2.TextRange.Characters.Text = CURL_COMMAND_TEXT
                    
                    #If Mac Then
                        Dim scriptResult As Boolean
                        scriptResult = AreAppleScriptsInstalled()
                    #Else
                        shps("Button_SpeakingEvalsScpt_Missing").Visible = True
                        shps("Button_DialogToolkit_Missing").Visible = True
                        shps("Button_EnhancedDialogs_Disable").Visible = True
                        shps("Button_SpeakingEvalsScpt_Installed").Visible = False
                        shps("Button_DialogToolkit_Installed").Visible = False
                        shps("Button_EnhancedDialogs_Enable").Visible = False
                    #End If
                    
                    Set shps = Nothing
                    SetLayoutMacOSUsers wb
                Case "mySignature"
                    SetLayoutMySignature wb
                Case Else
                    AutoPopulateEvaluationDateValues ws
                    SetLayoutClassRecords wb, ws
            End Select
            
            .Protect
            .EnableSelection = xlUnlockedCells
        End With
    Next ws
    
    startupMessageToDisplay = "Complete"
    
    #If Mac Then
        If Not scriptResult Then
            startupMessageToDisplay = "SpeakEvals.scpt Reminder"
            wb.Sheets("MacOS Users").Activate
        End If
    #End If
        
    DisplayStartupMessage startupMessageToDisplay

ReenableEvents:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub Workbook_SheetActivate(ByVal ws As Object)
    Const CURL_COMMAND_TEXT As String = "curl -L -o ~/Library/Application\ Scripts/com.microsoft.Excel/SpeakingEvals.scpt https://github.com/papercutter0324/SpeakingEvals/raw/main/SpeakingEvals.scpt"
    Dim wb As Workbook
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set wb = ThisWorkbook
    VerifySheetNames wb
    
    Select Case ws.Name
        Case "Instructions"
            SetLayoutInstructions wb, ws
        Case "MacOS Users"
            ws.Shapes("cURL_Command").TextFrame2.TextRange.Characters.Text = CURL_COMMAND_TEXT
            SetLayoutMacOSUsers wb
        Case "mySignature"
            SetLayoutMySignature wb
        Case Else
            If ws.Cells(1, 1).Value = "Native Teacher:" Then
                SetLayoutClassRecordsButtons ws
                ws.Cells(8, 2).Select
            End If
    End Select
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
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
            msgToDisplay = "Please wait while a self-check is performed and any errors " & vbNewLine & _
                           "are fixed. All existing data will be preserved." & vbNewLine & vbNewLine & _
                           "This should take less than a minute to complete."
            dialogSize = 430
        Case "Complete"
            msgToDisplay = "Process complete!"
            dialogSize = 120
        Case "SpeakEvals.scpt Reminder"
            msgToDisplay = "You must install SpeakingEvals.scpt for this file to fuction properly. Please follow the installation instructions and " & _
                           "read the notices about the System Events and File & Folder Permission requests."
            dialogSize = 470
    End Select
    
    msgResult = DisplayMessage(msgToDisplay, vbInformation, "Welcome!", dialogSize)
End Sub

Private Sub VerifySheetNames(ByRef wb As Workbook)
    Dim ws As Worksheet
    
    For Each ws In wb.Sheets
        Select Case ws.CodeName
            Case "Sheet1"
                If ws.Name <> "Instructions" Then ws.Name = "Instructions"
            Case "Sheet2"
                If ws.Name <> "MacOS Users" Then ws.Name = "MacOS Users"
            Case "Sheet3"
                If ws.Name <> "mySignature" Then ws.Name = "mySignature"
        End Select
    Next ws
End Sub

Private Sub AutoPopulateEvaluationDateValues(ByRef ws As Worksheet)
    Dim dateCell As Range, dateAsDate As Date, dateToCheck As Date
    Dim messageText As String, msgResult As Variant
    
    On Error Resume Next
    If ws.Range("A6").Value = "Evaluation Date:" Then
        Set dateCell = ws.Range("C6")

        If Len(Trim$(dateCell.Value)) = 0 Then
            dateCell.Value = Format(Date, "MMM. YYYY")
        ElseIf IsDate(Trim$(dateCell.Value)) Then
            dateAsDate = CDate(Trim$(dateCell.Value))
            dateToCheck = DateAdd("m", -2, Date)
        
            If dateAsDate < dateToCheck Then
                dateCell.Value = Format(Date, "MMM. YYYY")
            End If
        Else
            messageText = "An invalid date has been found on worksheet " & ws.Name & "." & vbNewLine & _
                          "Please enter a valid date."
            msgResult = DisplayMessage(messageText, vbInformation, "Invalid Date!")
            dateCell.Value = ""
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub SetLayoutInstructions(ByRef wb As Workbook, ByRef ws As Worksheet)
    Dim shp As Shapes, shapeProps As Variant, i As Integer
    Dim btnNamesArray As Variant
    
    Const TB_HEIGHT As Double = 58
    Const PADDING_LEFT As Double = 15
    Const PADDING_TOP As Double = 15
    Const SHAPE_SPACING As Double = 20
    
    Const INSTRUCTIONS_TB_TOP As Double = PADDING_TOP
    Const INSTRUCTIONS_TB_LEFT As Double = PADDING_LEFT
    Const INSTRUCTION_TB_HEIGHT As Double = TB_HEIGHT
    Const INSTRUCTIONS_TB_WIDTH As Double = 1165
    
    Const INSTRUCTIONS_MSG_TOP As Double = INSTRUCTIONS_TB_TOP + INSTRUCTION_TB_HEIGHT
    Const INSTRUCTIONS_MSG_LEFT As Double = INSTRUCTIONS_TB_LEFT
    Const INSTRUCTIONS_MSG_HEIGHT As Double = 560
    Const INSTRUCTIONS_MSG_WIDTH As Double = INSTRUCTIONS_TB_WIDTH
    
    Const CODE_TB_TOP As Double = INSTRUCTIONS_MSG_TOP + INSTRUCTIONS_MSG_HEIGHT + SHAPE_SPACING
    Const CODE_TB_LEFT As Double = PADDING_LEFT
    Const CODE_TB_HEIGHT As Double = TB_HEIGHT
    Const CODE_TB_WIDTH As Double = INSTRUCTIONS_MSG_WIDTH
    
    Const CODE_MSG_TOP As Double = CODE_TB_TOP + CODE_TB_HEIGHT
    Const CODE_MSG_LEFT As Double = CODE_TB_LEFT
    Const CODE_MSG_HEIGHT As Double = 265
    Const CODE_MSG_WIDTH As Double = CODE_TB_WIDTH
    
    Const WARNING_TB_TOP As Double = PADDING_TOP
    Const WARNING_TB_LEFT As Double = INSTRUCTIONS_TB_LEFT + INSTRUCTIONS_TB_WIDTH + SHAPE_SPACING
    Const WARNING_TB_HEIGHT As Double = TB_HEIGHT
    Const WARNING_TB_WIDTH As Double = 340
    
    Const WARNING_MSG_TOP As Double = WARNING_TB_TOP + WARNING_TB_HEIGHT
    Const WARNING_MSG_LEFT As Double = WARNING_TB_LEFT
    Const WARNING_MSG_HEIGHT As Double = 150
    Const WARNING_MSG_WIDTH As Double = WARNING_TB_WIDTH
    
    Const IMPORTANT_TB_TOP As Double = WARNING_MSG_TOP + WARNING_MSG_HEIGHT + SHAPE_SPACING
    Const IMPORTANT_TB_LEFT As Double = WARNING_MSG_LEFT
    Const IMPORTANT_TB_HEIGHT As Double = TB_HEIGHT
    Const IMPORTANT_TB_WIDTH As Double = WARNING_MSG_WIDTH
    
    Const IMPORTANT_MSG_TOP As Double = IMPORTANT_TB_TOP + IMPORTANT_TB_HEIGHT
    Const IMPORTANT_MSG_LEFT As Double = IMPORTANT_TB_LEFT
    Const IMPORTANT_MSG_HEIGHT As Double = 675
    Const IMPORTANT_MSG_WIDTH As Double = IMPORTANT_TB_WIDTH
    
    Const TODO_TB_TOP As Double = CODE_MSG_TOP + CODE_MSG_HEIGHT + SHAPE_SPACING
    Const TODO_TB_LEFT As Double = PADDING_LEFT
    Const TODO_TB_HEIGHT As Double = TB_HEIGHT
    Const TODO_TB_WIDTH As Double = INSTRUCTIONS_TB_WIDTH + SHAPE_SPACING + WARNING_TB_WIDTH
    
    Const TODO_MSG_TOP As Double = TODO_TB_TOP + TODO_TB_HEIGHT
    Const TODO_MSG_LEFT As Double = TODO_TB_LEFT
    Const TODO_MSG_HEIGHT As Double = 585
    Const TODO_MSG_WIDTH As Double = TODO_TB_WIDTH
    
    Const BUTTON_HEIGHT As Double = 70
    Const BUTTON_WIDTH As Double = 200
    Const BUTTON_TOP As Double = CODE_MSG_TOP + CODE_MSG_HEIGHT - BUTTON_HEIGHT - 20
    Const CELL_SPACING As Double = (CODE_MSG_WIDTH - 40 - (BUTTON_WIDTH * 5)) / 4
    
    btnNamesArray = Array("Button_Speadsheet", "Button_Font", "Button_ReportTemplate", "Button_SignatureTemplate", "Button_SourceCode")
    
    ' Define shape properties in an array: {Shape Name, Top, Left, Height, Width}
    shapeProps = Array( _
        Array("Title Bar - Instructions", INSTRUCTIONS_TB_TOP, INSTRUCTIONS_TB_LEFT, INSTRUCTION_TB_HEIGHT, INSTRUCTIONS_TB_WIDTH), _
        Array("Message - Instructions", INSTRUCTIONS_MSG_TOP, INSTRUCTIONS_MSG_LEFT, INSTRUCTIONS_MSG_HEIGHT, INSTRUCTIONS_MSG_WIDTH), _
        Array("Title Bar - Seeing the Code", CODE_TB_TOP, CODE_TB_LEFT, CODE_TB_HEIGHT, CODE_TB_WIDTH), _
        Array("Message - Seeing the Code", CODE_MSG_TOP, CODE_MSG_LEFT, CODE_MSG_HEIGHT, CODE_MSG_WIDTH), _
        Array("Title Bar - ToDo", TODO_TB_TOP, TODO_TB_LEFT, TODO_TB_HEIGHT, TODO_TB_WIDTH), _
        Array("Message - ToDo", TODO_MSG_TOP, TODO_MSG_LEFT, TODO_MSG_HEIGHT, TODO_MSG_WIDTH), _
        Array("Title Bar - Warning", WARNING_TB_TOP, WARNING_TB_LEFT, WARNING_TB_HEIGHT, WARNING_TB_WIDTH), _
        Array("Message - Warning", WARNING_MSG_TOP, WARNING_MSG_LEFT, WARNING_MSG_HEIGHT, WARNING_MSG_WIDTH), _
        Array("Title Bar - Important Files", IMPORTANT_TB_TOP, IMPORTANT_TB_LEFT, IMPORTANT_TB_HEIGHT, IMPORTANT_TB_WIDTH), _
        Array("Message - Important Files", IMPORTANT_MSG_TOP, IMPORTANT_MSG_LEFT, IMPORTANT_MSG_HEIGHT, IMPORTANT_MSG_WIDTH) _
    )
    
    Set shp = ws.Shapes
    
    ' Loop through the shape properties array and apply the settings
    For i = LBound(shapeProps) To UBound(shapeProps)
        With shp.Item(shapeProps(i)(0))
            .Top = shapeProps(i)(1)
            .Left = shapeProps(i)(2)
            .Height = shapeProps(i)(3)
            .Width = shapeProps(i)(4)
        End With
    Next i
    
    ' Position buttons
    For i = LBound(btnNamesArray) To UBound(btnNamesArray)
        With shp.Item(btnNamesArray(i))
            .Top = BUTTON_TOP
            .Left = CODE_MSG_LEFT + 20 + i * (BUTTON_WIDTH + CELL_SPACING)
            .Height = BUTTON_HEIGHT
            .Width = BUTTON_WIDTH
        End With
    Next i
End Sub

Private Sub SetLayoutMacOSUsers(ByRef wb As Workbook)
    Dim shp As Shapes
    Dim shapeProps As Variant, buttonProps As Variant
    Dim i As Integer
    
    Const MACOS_TB_TOP As Double = 15
    Const MACOS_TB_LEFT As Double = 15
    Const MACOS_TB_HEIGHT As Double = 58
    Const MACOS_TB_WIDTH As Double = 1285
    
    Const MACOS_MSG_TOP As Double = MACOS_TB_TOP + MACOS_TB_HEIGHT
    Const MACOS_MSG_LEFT As Double = MACOS_TB_LEFT
    Const MACOS_MSG_HEIGHT As Double = 710
    Const MACOS_MSG_WIDTH As Double = MACOS_TB_WIDTH
    
    Const CURL_TOP As Double = MACOS_MSG_TOP
    Const CURL_HEIGHT As Double = 115
    Const CURL_WIDTH As Double = 560
    Const CURL_LEFT As Double = MACOS_MSG_LEFT + MACOS_MSG_WIDTH - CURL_WIDTH
    
    Const BUTTON_HEIGHT As Double = 70
    Const BUTTON_WIDTH As Double = 200
    Const BUTTON_TOP As Double = MACOS_MSG_TOP + MACOS_MSG_HEIGHT - BUTTON_HEIGHT - 20
    
    ' Define shape properties in an array: {Shape Name, Top, Left, Height, Width}
    shapeProps = Array( _
        Array("Title Bar", MACOS_TB_TOP, MACOS_TB_LEFT, MACOS_TB_HEIGHT, MACOS_TB_WIDTH), _
        Array("Message", MACOS_MSG_TOP, MACOS_MSG_LEFT, MACOS_MSG_HEIGHT, MACOS_MSG_WIDTH), _
        Array("cURL_Command", CURL_TOP, CURL_LEFT, CURL_HEIGHT, CURL_WIDTH) _
    )
    
    ' Define button properties in an array: {Button Name, Left Position}
    buttonProps = Array( _
        Array("Button_SpeakingEvalsScpt_Installed", MACOS_MSG_LEFT + 70), _
        Array("Button_SpeakingEvalsScpt_Missing", MACOS_MSG_LEFT + 70), _
        Array("Button_DialogToolkit_Installed", MACOS_MSG_LEFT + 350), _
        Array("Button_DialogToolkit_Missing", MACOS_MSG_LEFT + 350), _
        Array("Button_EnhancedDialogs_Enable", MACOS_MSG_LEFT + 630), _
        Array("Button_EnhancedDialogs_Disable", MACOS_MSG_LEFT + 630) _
    )
    
    Set shp = wb.Sheets("MacOS Users").Shapes
    
    ' Loop through the shape properties array and apply the settings
    For i = LBound(shapeProps) To UBound(shapeProps)
        With shp.Item(shapeProps(i)(0))
            .Top = shapeProps(i)(1)
            .Left = shapeProps(i)(2)
            .Height = shapeProps(i)(3)
            .Width = shapeProps(i)(4)
        End With
    Next i

    ' Loop through button properties and set positions
    For i = LBound(buttonProps) To UBound(buttonProps)
        With shp.Item(buttonProps(i)(0))
            .Top = BUTTON_TOP
            .Left = buttonProps(i)(1)
            .Height = BUTTON_HEIGHT
            .Width = BUTTON_WIDTH
        End With
    Next i
End Sub

Private Sub SetLayoutMySignature(ByRef wb As Workbook)
    Dim shp As Shapes
    Dim shapeProps As Variant
    Dim maxHeight As Double, maxWidth As Double, aspectRatio As Double
    Dim i As Integer
    
    Const TB_HEIGHT As Double = 58
    Const TB_WIDTH As Double = 1270
    Const TB_TOP As Double = 15
    Const TB_LEFT As Double = 15
    
    Const MSG_HEIGHT As Double = 640
    Const SIG_TB_WIDTH As Double = 300
    Const SIG_CONTAINER_HEIGHT As Double = 86
    
    maxHeight = 68.2
    maxWidth = 286

    ' Define shape properties in an array: {Shape Name, Top, Left, Height, Width}
    shapeProps = Array( _
        Array("Title Bar", TB_TOP, TB_LEFT, TB_HEIGHT, TB_WIDTH), _
        Array("Message", TB_TOP + TB_HEIGHT, TB_LEFT, MSG_HEIGHT, TB_WIDTH), _
        Array("Signature Title Bar", TB_TOP, TB_LEFT + TB_WIDTH - SIG_TB_WIDTH, TB_HEIGHT, SIG_TB_WIDTH), _
        Array("Signature Container", TB_TOP + TB_HEIGHT, TB_LEFT + TB_WIDTH - SIG_TB_WIDTH, SIG_CONTAINER_HEIGHT, SIG_TB_WIDTH) _
    )

    Set shp = wb.Sheets("mySignature").Shapes

    ' Loop through shape properties array and apply settings
    For i = LBound(shapeProps) To UBound(shapeProps)
        With shp.Item(shapeProps(i)(0))
            .Top = shapeProps(i)(1)
            .Left = shapeProps(i)(2)
            .Height = shapeProps(i)(3)
            .Width = shapeProps(i)(4)
        End With
    Next i

    ' Center signature images if they exist
    If DoesShapeExist(wb.Sheets("mySignature"), "mySignature_Placeholder") Then
        With shp.Item("mySignature_Placeholder")
            .LockAspectRatio = msoFalse
            .Height = maxHeight
            .Width = maxWidth
            .LockAspectRatio = msoTrue
            .Top = shp.Item("Signature Container").Top + (SIG_CONTAINER_HEIGHT / 2) - (.Height / 2)
            .Left = shp.Item("Signature Container").Left + (SIG_TB_WIDTH / 2) - (.Width / 2)
        End With
    End If

    If DoesShapeExist(wb.Sheets("mySignature"), "mySignature") Then
        With shp.Item("mySignature")
            aspectRatio = .Width / .Height
            
            If maxWidth / maxHeight > aspectRatio Then
                .Width = maxHeight * aspectRatio
                .Height = maxHeight
            Else
                .Width = maxWidth
                .Height = maxWidth / aspectRatio
            End If

            .Top = shp.Item("Signature Container").Top + (SIG_CONTAINER_HEIGHT / 2) - (.Height / 2)
            .Left = shp.Item("Signature Container").Left + (SIG_TB_WIDTH / 2) - (.Width / 2)
        End With
    End If
End Sub

Private Function DoesShapeExist(ByVal ws As Worksheet, ByVal shapeName As String) As Boolean
    On Error Resume Next
    DoesShapeExist = Not ws.Shapes(shapeName) Is Nothing
    On Error GoTo 0
End Function

Private Sub SetLayoutClassRecords(ByRef wb As Workbook, ByRef ws As Worksheet)
    Dim cellTop As Double, cellHeight As Double, cellLeft As Double, cellWidth As Double
    Dim shadingRanges As Variant
    Dim currentRng As Range
    Dim i As Integer
    
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
    
    With wb.Names
        For i = LBound(shadingRanges) To UBound(shadingRanges)
            Set currentRng = ws.Range(shadingRanges(i)(1))
            On Error Resume Next
            If Not .Item(shadingRanges(i)(0)) Is Nothing Then .Item(shadingRanges(i)(0)).Delete
            On Error GoTo 0
            .Add Name:=shadingRanges(i)(0), RefersTo:=currentRng
            currentRng.Interior.Color = shadingRanges(i)(2)
            VerifyValidationSettings ws, shadingRanges(i)(0), currentRng
        Next i
    End With

    With ws
        With .Columns
            .Item("A").ColumnWidth = 7
            .Item("B:C").ColumnWidth = 18
            .Item("D:I").ColumnWidth = 21
            .Item("J").ColumnWidth = 102.5
            .Item("K").ColumnWidth = 44.17
        End With

        With .Rows
            .Item("1:6").RowHeight = 30
            .Item("7").RowHeight = 25
            .Item("8:32").RowHeight = 50
        End With
        
        SetLayoutClassRecordsButtons ws
        
        ' Set thick borders
        With .Range("A1:C6,A8:A32,B8:C32,D8:I32,J8:J32,K8:K32").Borders
            .LineStyle = xlContinuous
            .Weight = xlThick
        End With
            
        ' Set inside dashed borders
        With .Range("A1:C6,A8:A32,B8:C32,D8:I32,J8:J32,K8:K32").Borders
            .Item(xlInsideHorizontal).LineStyle = xlDash
            .Item(xlInsideHorizontal).Weight = xlThin
            .Item(xlInsideVertical).LineStyle = xlLineStyleNone
        End With
            
        ' Set inside no borders
        With .Range("D1:I6,J1:J5,K1:K6").Borders
            .Item(xlInsideHorizontal).LineStyle = xlLineStyleNone
            .Item(xlInsideVertical).LineStyle = xlLineStyleNone
        End With
            
        ' Set unlocked cells, font alignment, and text formatting
        With .Range("C1:C6,B8:K32")
            .Locked = False
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            .WrapText = True
            .NumberFormat = "@"
        End With
            
        ' Set locked cells
        .Range("A1:B6,D1:K6,A7:K7,A8:A32").Locked = True
    End With
End Sub

Private Sub SetLayoutClassRecordsButtons(ByRef ws As Worksheet)
    Dim cellTop As Double, cellHeight As Double, cellLeft As Double, cellWidth As Double
    Dim cellVerticalSpacing As Double, cellHorizontalSpacing As Double
    Dim buttonProps As Variant
    Dim i As Integer
    
    Const BUTTON_HEIGHT As Double = 50
    Const BUTTON_WIDTH As Double = 187
    
    With ws.Cells(1, 10)
        cellTop = .Top
        cellHeight = .Height * 5
        cellLeft = .Left
        cellWidth = .Width
    End With
    
    cellVerticalSpacing = (cellHeight - (2 * BUTTON_HEIGHT)) / 3
    cellHorizontalSpacing = (cellWidth - (2 * BUTTON_WIDTH)) / 3
    
    ' Define button properties in an array: {Button Name, Row Index, Col Index}
    buttonProps = Array( _
        Array("Button_GenerateProofs", 1, 1), _
        Array("Button_GenerateReports", 2, 1), _
        Array("Button_OpenTypingSite", 1, 2), _
        Array("Button_RepairLayout", 2, 2) _
    )

    ' Loop through button array and set positions
    With ws.Shapes
        For i = LBound(buttonProps) To UBound(buttonProps)
            With .Item(buttonProps(i)(0))
                .Height = BUTTON_HEIGHT
                .Width = BUTTON_WIDTH
                .Top = cellTop + (cellVerticalSpacing * buttonProps(i)(1)) + ((buttonProps(i)(1) - 1) * BUTTON_HEIGHT)
                .Left = cellLeft + (cellHorizontalSpacing * buttonProps(i)(2)) + ((buttonProps(i)(2) - 1) * BUTTON_WIDTH)
            End With
        Next i
    End With
End Sub

Private Sub VerifyValidationSettings(ByRef ws As Worksheet, ByVal rangeType As String, ByRef cellRange As Range)
    Dim currentCell As Range, rangeCol As Range, rangeCell As Range
    Dim dateInputMessage As Variant, validationValues As Variant
    Dim validationInputTitle As String, validationInputMessage As String
    Dim columnFont As String, columnFontSize As Integer, columnFontBold As Boolean
    Dim i As Double
    
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
                Array("Native Teacher's Name", _
                      "Please enter just your" & vbNewLine & "name, no suffix or title" & vbNewLine & "like ""tr.""", _
                      ""), _
                Array("Korean Teacher's Name", _
                      "Please write their Korean name. The parents are unlikely to know their English name.", _
                      ""), _
                Array("Class Level", _
                      "Click on the down arrow and choose the class's level from the list.", _
                      "Theseus, Perseus, Odysseus, Hercules, Artemis, Hermes, Apollo, Zeus, E5 Athena, Helios, Poseidon, Gaia, Hera, E6 Song's"), _
                Array("Class Days", _
                      "Select the days when you see this class." & vbNewLine & vbNewLine & "For Athena and Song's classes, use Class-1 and Class-2 to help organize split classes.", _
                      "MonWed, MonFri, WedFri, MWF, TTh, MWF (Class 1), MWF (Class 2), TTh (Class 1), TTh (Class 2)"), _
                Array("Class Time", _
                      "Select what time you have Class 1 each week. Scroll to see more options." & vbNewLine & vbNewLine & "This is to help you keep track of which class this is; it won't appear on the final reports.", _
                      "9pm, 830pm, 8pm, 7pm, 6pm, 530pm, 5pm, 4pm, 3pm, 2pm, 1pm, 12pm, 11am, 10am, 9am"), _
                Array("Date Format", _
                      dateInputMessage, _
                      "") _
            )
            
            For i = 1 To 6
                columnFont = IIf(i = 2, "Batang", "Calibri")
                SetCellFormating ws.Cells(i, 3), validationValues(i - 1)(0), validationValues(i - 1)(1), validationValues(i - 1)(2), columnFont, 14, False
            Next i
        Case "EngKorNameShadingRange", "grammarValuesShadingRange", "pronunciationValuesShadingRange", "fluencyValuesShadingRange", "mannerValuesShadingRange", "contentValuesShadingRange", "overallEffortValuesShadingRange", "commentValuesShadingRange", "notesValuesShadingRange"
            columnFont = "Calibri"
            columnFontBold = False
            
            For Each rangeCol In cellRange.Columns
                Select Case rangeCol.Column
                    Case 2
                        validationInputTitle = "Character Limit"
                        validationInputMessage = "30 characters"
                        columnFontSize = 18
                    Case 3
                        validationInputTitle = "Language Reminder"
                        validationInputMessage = "Please write their names in Korean."
                        columnFont = "Batang"
                        columnFontSize = 20
                    Case 4 To 9
                        validationInputTitle = "Enter a Grade"
                        validationInputMessage = "Valid Letter Grades" & vbNewLine & "  A+ / A / B+ / B / C" & vbNewLine & vbNewLine & "Valid Numeric Scores   " & vbNewLine & "  1 ~ 5"
                        columnFontSize = 22
                        columnFontBold = True
                    Case 10
                        validationInputTitle = "Character Limit"
                        validationInputMessage = "315 characters"
                        columnFontSize = 14
                    Case 11
                        validationInputTitle = ""
                        validationInputMessage = ""
                        columnFontSize = 14
                End Select
                
                For Each rangeCell In rangeCol.Cells
                    SetCellFormating rangeCell, validationInputTitle, validationInputMessage, "", columnFont, columnFontSize, columnFontBold
                Next rangeCell
            Next rangeCol
    End Select
End Sub

Private Sub SetCellFormating(ByRef wsCell As Range, ByVal validationInputTitle As String, ByVal validationInputMessage As String, ByVal validationListValue As String, ByVal columnFont As String, ByVal columnFontSize As Integer, ByVal columnFontBold As Boolean)
    Dim cellRow As Long, cellColumn As Long

    On Error Resume Next
    With wsCell
        cellRow = .Row
        cellColumn = .Column
        
        With .Validation
            If 1 = 1 Then '.InputTitle <> validationInputTitle Then
                .Delete
                
                If cellColumn = 3 Then
                    Select Case cellRow
                        Case 3 To 5
                            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=validationListValue
                            .IgnoreBlank = True
                            .InCellDropdown = True
                            .ShowInput = True
                            .ShowError = True
                        Case Else
                            .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop
                    End Select
                Else
                    .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop
                End If

                .InputTitle = validationInputTitle
                .InputMessage = validationInputMessage
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
' Report Generation
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GenerateReports(ByRef ws As Worksheet, ByVal clickedButtonName As String)
    Const REPORT_TEMPLATE As String = "SpeakingEvaluationTemplate.pptx"
    Const ERR_RESOURCES_FOLDER As String = "resourcesFolder"
    Const ERR_INCOMPLETE_RECORDS As String = "incompleteRecords"
    Const ERR_LOADING_POWERPOINT As String = "loadingPowerPoint"
    Const ERR_LOADING_TEMPLATE As String = "loadingTemplate"
    Const ERR_MISSING_SHAPES As String = "missingTemplateShapes"
    Const MSG_SAVE_FAILED As String = "exportFailed"
    Const MSG_ZIP_FAILED As String = "zipFailed"
    Const MSG_SUCCESS As String = "exportSuccessful"

    ' Objects to open PowerPoint and modify the template
    Dim pptApp As Object, pptDoc As Object
    
    ' Variables for determining the code path and important states
    Dim generateProcess As String, saveResult As Boolean
    
    ' Strings for generating important messages for the user
    Dim resultMsg As String, msgToDisplay As String, msgTitle As String
    Dim msgType As Integer, msgResult As Variant, dialogSize As Integer
    
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
        If Not RequestFileAndFolderAccess(resourcesFolder) Then
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
    
    If Not InstallFonts() Then
        ' Throw an error
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Font required for reports not installed. Please install manually before continuing."
        #End If
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
    If templatePath = "" Then
        ' Set an error msg
        GoTo CleanUp
    End If

    savePath = SetSaveLocation(ws, generateProcess, resourcesFolder)
    If savePath = "" Then
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
        Debug.Print "Beginning to generate reports."
    #End If
    
    For currentRow = firstStudentRecord To lastRow
        #If PRINT_DEBUG_MESSAGES Then
            i = i + 1
            Debug.Print "    Current report: " & i & " of " & (lastRow - firstStudentRecord + 1)
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
        Case ERR_LOADING_POWERPOINT, ERR_LOADING_TEMPLATE
            msgToDisplay = "There was an error opening MS PowerPoint and/or the template. This is sometimes normal MS Office behaviour, so please wait a couple seconds and try again."
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
    If Not pptApp Is Nothing Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Beginning final cleanup checks."
        #End If
        KillPowerPoint pptApp, pptDoc
    End If
End Sub

Private Function LoadPowerPoint(ByRef pptApp As Object, ByRef pptDoc As Object, ByVal templatePath As String) As Boolean
    Dim openDoc As Object
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Attempting to open an instance of MS PowerPoint."
    #End If
    
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    Err.Clear
    On Error GoTo ErrorHandler
    
    ' Open a new instance of PowerPoint if needed
    #If Mac Then
        Dim appleScriptResult As String, msgToDisplay As String, msgResult As Variant
        
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
        Debug.Print "    MS PowerPoint loaded: " & (Not pptApp Is Nothing)
    #End If
    
    ' Make the process visible so users understand their computer isn't frozen
    With pptApp
        .Visible = True
    End With
    
    If Not pptApp Is Nothing Then
        Set pptDoc = pptApp.Presentations.Open(templatePath)
        If Val(pptApp.Version) > 15 Then
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "    Attempting to disable AutoSave."
            #End If
            DisableAutoSave pptDoc
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "    AutoSave status: " & pptDoc.AutoSaveOn
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
        msgResult = DisplayMessage(msgToDisplay, vbOKOnly, "Error Loading PowerPoint", 470)
    #End If
    LoadPowerPoint = False
End Function

Private Sub WritePptReport(ByRef ws As Object, ByRef pptApp As Object, ByRef pptDoc As Object, ByVal generateProcess As String, ByVal currentRow As Integer, ByVal savePath As String, ByRef saveResult As Boolean)
    Dim englishName As String, koreanName As String, classLevel As String, nativeTeacher As String, koreanTeacher As String, evalDate As String
    Dim commentText As String, classTime As String, fileName As String, validEnglishName As String
    Dim scoreCategories As Variant, scoreValues As Variant
    Dim i As Integer
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "        Preparing report data."
    #End If
    
    With ws.Cells
        ' Header values
        englishName = .Item(currentRow, 2).Value
        koreanName = .Item(currentRow, 3).Value
        classLevel = .Item(3, 3).Value
        nativeTeacher = .Item(1, 3).Value
        koreanTeacher = .Item(2, 3).Value
        evalDate = Format(.Item(6, 3).Value, "MMM. YYYY")

        ' Scores and comment values
        scoreCategories = Array("Grammar_", "Pronunciation_", "Fluency_", "Manner_", "Content_", "Effort_", "Result_")
        scoreValues = Array(.Item(currentRow, 4).Value, .Item(currentRow, 5).Value, .Item(currentRow, 6).Value, _
                           .Item(currentRow, 7).Value, .Item(currentRow, 8).Value, .Item(currentRow, 9).Value, _
                           CalculateOverallGrade(ws, currentRow))
        commentText = .Item(currentRow, 10).Value

        ' Set report's filename
        classTime = .Item(4, 3).Value & "-" & .Item(5, 3).Value
        
        validEnglishName = SanitizeFileName(englishName)
        
        fileName = koreanName & "(" & validEnglishName & ")" & " - " & .Item(4, 3).Value
    End With
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "        Report filename: " & fileName & vbNewLine & _
                    "        Saving to: " & savePath
    #End If
    
    On Error Resume Next
    With pptDoc.Slides(1).Shapes
        With .Item("Report_Header").GroupItems
             .Item("English_Name").TextFrame.TextRange.Text = englishName
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

        On Error Resume Next
        If .Item("Signature") Is Nothing Then InsertSignature pptDoc
        On Error GoTo 0
    End With
    On Error GoTo 0
    
    saveResult = SavePptToFile(pptApp, pptDoc, generateProcess, savePath, fileName)
End Sub

Private Function CalculateOverallGrade(ByRef ws As Worksheet, ByVal currentRow As Integer) As String
    Dim scoreRange As Range, gradeCell As Range
    Dim totalScore As Double, avgScore As Double
    Dim roundedScore As Integer, numericScore As Integer
    
    Set scoreRange = ws.Range("D" & currentRow & ":" & "I" & currentRow)
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

Private Function SanitizeFileName(ByVal englishName As String) As String
    Dim invalidCharacters As Variant, reservedNames As Variant, ch As Variant
    
    invalidCharacters = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    
    For Each ch In invalidCharacters
        englishName = Replace(englishName, ch, "_")
    Next ch
    
    englishName = Trim(englishName)
    
    Do While Right(englishName, 1) = "."
        englishName = Left(englishName, Len(englishName) - 1)
    Loop
    
    If Len(englishName) > 10 Then englishName = Trim(Left(englishName, 10))
    
    Do While Right(englishName, 1) = "_"
        englishName = Left(englishName, Len(englishName) - 1)
    Loop
    
    SanitizeFileName = Trim(englishName)
End Function

Private Sub ToggleScoreVisibility(ByRef pptDoc As Object, ByVal scoreCategory As String, ByVal scoreValue As String)
    With pptDoc.Slides(1).Shapes(scoreCategory & "Scores").GroupItems
        .Item(scoreCategory & "A+").Visible = IIf(scoreValue = "A+", msoTrue, msoFalse)
        .Item(scoreCategory & "A").Visible = IIf(scoreValue = "A", msoTrue, msoFalse)
        .Item(scoreCategory & "B+").Visible = IIf(scoreValue = "B+", msoTrue, msoFalse)
        .Item(scoreCategory & "B").Visible = IIf(scoreValue = "B", msoTrue, msoFalse)
        .Item(scoreCategory & "C").Visible = IIf(scoreValue = "C", msoTrue, msoFalse)
    End With
End Sub

Private Sub InsertSignature(ByRef pptDoc As Object)
    Dim sigShape As Object, sigWidth As Double, sigHeight As Double, sigAspectRatio As Double
    Dim signatureFound As Boolean
    
    Const SIGNATURE_SHAPE_NAME As String = "mySignature"
    
    ' These numbers make no sense, but they work.
    Const ABSOLUTE_LEFT As Double = 375
    Const ABSOLUTE_TOP As Double = 727.5
    Const MAX_WIDTH As Double = 130
    Const MAX_HEIGHT As Double = 31
    
    Static signaturePath As String
    Static signatureImagePath As String
    Static useEmbeddedSignature As Boolean
    
    If signaturePath = "" Then
        signaturePath = ThisWorkbook.Path & Application.PathSeparator
        ConvertOneDriveToLocalPath signaturePath
    End If
    
    On Error Resume Next
    Set sigShape = pptDoc.Slides(1).Shapes(SIGNATURE_SHAPE_NAME)
    If Not sigShape Is Nothing Then Exit Sub
    useEmbeddedSignature = (Not ThisWorkbook.Sheets("mySignature").Shapes(SIGNATURE_SHAPE_NAME) Is Nothing)
     
    If useEmbeddedSignature Then
        ExportSignatureFromExcel SIGNATURE_SHAPE_NAME, signatureImagePath
    Else
        signatureImagePath = GetSignatureFile(signaturePath)
        If signatureImagePath = "" Then Exit Sub
    End If
    
    Set sigShape = pptDoc.Slides(1).Shapes.AddPicture(fileName:=signatureImagePath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                                                      Left:=ABSOLUTE_LEFT, Top:=ABSOLUTE_TOP)
    If Err.Number <> 0 Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "Error inserting ignature."
        #End If
        Exit Sub
    End If
    On Error GoTo 0
    
    sigShape.Name = SIGNATURE_SHAPE_NAME
    
    ' Maintain the aspect ratio and resize if needed
    sigAspectRatio = sigShape.Width / sigShape.Height
    If MAX_WIDTH / MAX_HEIGHT > sigAspectRatio Then
        ' Adjust width to fit within max height
        sigWidth = MAX_HEIGHT * sigAspectRatio
        sigHeight = MAX_HEIGHT
    Else
        ' Adjust height to fit within max width
        sigWidth = MAX_WIDTH
        sigHeight = MAX_WIDTH / sigAspectRatio
    End If

    ' Position and resize the image
    With sigShape
        .LockAspectRatio = msoTrue
        .Width = sigWidth
        .Height = sigHeight
    End With
End Sub

Private Sub ExportSignatureFromExcel(ByVal SIGNATURE_SHAPE_NAME As String, signatureImagePath As String)
    Dim signSheet As Worksheet, tempSheet As Worksheet, signatureshp As Shape, chrtObj As ChartObject
    
    Application.DisplayAlerts = False
    
    Set tempSheet = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
    tempSheet.Name = "Temp_signature"
    
    Set signatureshp = ThisWorkbook.Worksheets("mySignature").Shapes(SIGNATURE_SHAPE_NAME)
    signatureshp.Copy
    
    signatureImagePath = GetTempFilePath("tempSignature.png")
    ConvertOneDriveToLocalPath signatureImagePath
    
    On Error Resume Next
    Kill signatureImagePath
    Err.Clear
    
    Set chrtObj = tempSheet.ChartObjects.Add(Left:=tempSheet.Range("B2").Left, _
                                             Top:=tempSheet.Range("B2").Top, _
                                             Width:=signatureshp.Width, _
                                             Height:=signatureshp.Height)
    With chrtObj
        .Activate
        DoEvents
        .Chart.Paste
        Application.Wait Now + TimeValue("00:00:01")
        .Chart.ChartArea.Format.Line.Visible = msoFalse
        DoEvents
        .Chart.Export signatureImagePath, "png"
        DoEvents
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
        If Dir(signaturePath & "mySignature.png") <> "" Then
            GetSignatureFile = signaturePath & "mySignature.png"
        ElseIf Dir(signaturePath & "mySignature.jpg") <> "" Then
            GetSignatureFile = signaturePath & "mySignature.jpg"
        Else
            GetSignatureFile = ""
        End If
    #End If
End Function

Private Function SavePptToFile(ByRef pptApp As Object, ByRef pptDoc As Object, ByVal saveRoutine As String, ByVal savePath As String, ByVal fileName As String) As Boolean
    Dim tempFile As String, destFile As String
    
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
    
    If Val(pptApp.Version) > 15 Then DisableAutoSave pptDoc
    
    SavePptToFile = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub DisableAutoSave(ByRef pptDoc As Object)
    On Error Resume Next
    If pptDoc.AutoSaveOn Then pptDoc.AutoSaveOn = False
    On Error GoTo 0
End Sub

Private Sub KillPowerPoint(ByRef pptApp As Object, ByRef pptDoc As Object)
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Attempting to close the open instance of MS PowerPoint."
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
            Debug.Print "    Attempting to close MS PowerPoint."
        #End If
        pptApp.Quit
        Set pptApp = Nothing
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Status: " & (pptApp Is Nothing)
        #End If
    End If

    #If Mac Then
        Dim closeResult As String
        
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

Private Sub ZipReports(ByRef ws As Worksheet, ByVal savePath As Variant, ByRef saveResult As Boolean, ByVal resourcesFolder As String)
    Dim zipCommand As String, zipPath As Variant, zipName As Variant, pdfPath As Variant
    Dim classLevel As String, classKT As String, classDays As String
    Dim errDescription As String
    Dim archiverPath As String
    
    On Error Resume Next
    If Right(savePath, 1) <> Application.PathSeparator Then savePath = savePath & Application.PathSeparator
    
    With ws.Cells
        classLevel = .Item(3, 3).Value
        classKT = .Item(2, 3).Value
        classDays = .Item(4, 3).Value
    End With
    
    zipName = classLevel & " (" & classKT & " - " & classDays & ").zip"
    
    #If Mac Then
        zipPath = savePath & zipName
    #Else
        Dim fso As Object
        
        zipPath = GetTempFilePath(zipName)
        
        ' Remove old copy if present
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True
    #End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Attempting to create a ZIP file of all generated reports for this class." & vbNewLine & _
                    "    Filename: " & zipName & vbNewLine & _
                    "    Path: " & savePath
    #End If
    
    #If Mac Then
        Dim scriptResultString As String, scriptResultBoolean As Boolean
        
        archiverPath = FindPathToArchiveTool(resourcesFolder)
            
        Select Case archiverPath
            Case ""
                scriptResultString = AppleScriptTask(APPLE_SCRIPT_FILE, "CreateZipWithDefaultArchiver", savePath & APPLE_SCRIPT_SPLIT_KEY & zipPath)
            Case Else
                zipCommand = Chr(34) & archiverPath & Chr(34) & " a " & Chr(34) & zipPath & Chr(34) & " " & Chr(34) & savePath & "*.pdf" & Chr(34)
                scriptResultString = AppleScriptTask(APPLE_SCRIPT_FILE, "CreateZipWithLocal7Zip", zipCommand)
        End Select
        
        If scriptResultString <> "Success" Then
            errDescription = scriptResultString
            saveResult = False
        Else
            saveResult = True
            scriptResultBoolean = AppleScriptTask(APPLE_SCRIPT_FILE, "ClearPDFsAfterZipping", savePath)
        End If
    #Else
        Dim shellApp As Object, archiverName As String, startTime As Double
        startTime = Timer
        
        archiverPath = FindPathToArchiveTool(resourcesFolder, archiverName)
        
        Select Case archiverName
            Case "7Zip", "Local 7zip"
                zipCommand = Chr(34) & archiverPath & Chr(34) & " a " & Chr(34) & zipPath & Chr(34) & " " & Chr(34) & savePath & "*.pdf" & Chr(34)
                Shell zipCommand, vbNormalFocus
            Case Else
                Set shellApp = CreateObject("Shell.Application")
                
                ' Simplify the filename in case Hangul in filenames isn't fully supported
                zipName = classLevel & " (" & classDays & ").zip"
                zipPath = GetTempFilePath(zipName)
        
                ' Retry removing old copy if present
                If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True
                
                ' Create an empty ZIP file
                Open zipPath For Output As #1
                Print #1, "PK" & Chr(5) & Chr(6) & String(18, vbNullChar)
                Close #1
                                
                ' Add the contents of savePath to the zip file
                shellApp.Namespace(zipPath).CopyHere shellApp.Namespace(savePath).Items
        End Select
        
        Do ' Wait for the zip file to be created, but no longer than 10 seconds
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop While Not fso.FileExists(zipPath) And Timer - startTime < 10
        
        ' Copy the zip file and report if process was successful
        If fso.FileExists(zipPath) Then
            Application.Wait (Now + TimeValue("0:00:02")) ' Wait a couple seconds for the file to be released
            fso.CopyFile zipPath, savePath & zipName, True
            Kill zipPath
            saveResult = True
        Else
            If Err.Number <> 0 Then errDescription = Err.Description
            saveResult = False
        End If
        
        ' Clear out PDFs
        If saveResult Then DeletePDFs savePath
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

Private Sub DeletePDFs(ByVal targetFolder As String)
    #If Mac Then
    #Else
        Dim fso As Object, objFile As Object, objFolder As Object
    
        If Right(targetFolder, 1) <> Application.PathSeparator Then targetFolder = targetFolder & Application.PathSeparator
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set objFolder = fso.GetFolder(targetFolder)
        
        On Error Resume Next
        For Each objFile In objFolder.Files
            If LCase(fso.GetExtensionName(objFile.Name)) = "pdf" Then
                objFile.Delete True
                #If PRINT_DEBUG_MESSAGES Then
                    If Err.Number <> 0 Then
                        Debug.Print "Error deleting: " & objFile.Name & vbNewLine & _
                                    "Error: " & Err.Description
                        Err.Clear
                    End If
                #End If
            End If
        Next objFile
        On Error GoTo 0
    #End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Data Validation
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
    lastRow = ws.Cells(ws.Rows.Count, STUDENT_INFO_FIRST_COL).End(xlUp).Row
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
    
    dataValue = Trim$(currentCell.Value)
    
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
            ValidateData = (Len(dataValue) < 960)
        Case Else
            ValidateData = False
    End Select
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
' Resources Management
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Download7Zip(ByVal resourcesFolder As String, ByRef downloadResult As Boolean)
    Dim destinationPath As String, downloadURL As String
    
    Const GIT_REPO_URL As String = "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/"
    
    #If Mac Then
        Dim scriptResultBoolean As Boolean
        
        Const FILE_NAME As String = "7zz"
        
        destinationPath = resourcesFolder & Application.PathSeparator & FILE_NAME
        downloadURL = GIT_REPO_URL & FILE_NAME
        
        scriptResultBoolean = AppleScriptTask(APPLE_SCRIPT_FILE, "DownloadFile", destinationPath & APPLE_SCRIPT_SPLIT_KEY & downloadURL)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print IIf(scriptResultBoolean, "    Download successful.", "    Error: " & Err.Description)
        #End If
        
        If scriptResultBoolean Then
            downloadResult = RequestFileAndFolderAccess(resourcesFolder, destinationPath)
            scriptResultBoolean = AppleScriptTask(APPLE_SCRIPT_FILE, "ChangeFilePermissions", "+x" & APPLE_SCRIPT_SPLIT_KEY & destinationPath)
        End If
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    File access " & IIf(downloadResult, "granted.", "denied.")
        #End If
    #Else
        Dim objWMI As Object, colProcessors As Object, objProcessor As Object
        Dim architecture As String, fileToDownload As String
        
        Const FILE_NAME As String = "7za.exe"
        
        Set objWMI = GetObject("winmgmts:\\.\root\CIMV2")
        Set colProcessors = objWMI.ExecQuery("SELECT Architecture FROM Win32_Processor")
        
        For Each objProcessor In colProcessors
            Select Case objProcessor.architecture
                Case 0: fileToDownload = "7za(x86).exe"
                Case 9: fileToDownload = "7za(x64).exe"
                Case 12: fileToDownload = "7za(ARM).exe"
            End Select
        Next
        
        destinationPath = resourcesFolder & Application.PathSeparator & FILE_NAME
        downloadURL = GIT_REPO_URL & fileToDownload
        
        If Dir(destinationPath) <> "" Then
            ' Add a hash check to verify the file
            downloadResult = True
            Exit Sub
        End If
        
        Select Case True
            Case CheckForCurl()
                downloadResult = DownloadUsingCurl(destinationPath, downloadURL)
            Case CheckForDotNet35()
                downloadResult = DownloadUsingDotNet35(destinationPath, downloadURL)
            Case Else
                downloadResult = False
        End Select
    #End If
End Sub

Private Function DownloadReportTemplate(ByVal templatePath As String, ByVal resourcesFolder As String) As Boolean
    Const REPORT_TEMPLATE_URL As String = "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/SpeakingEvaluationTemplate.pptx"
    Dim downloadResult As Boolean
    
    #If Mac Then
        On Error Resume Next
        downloadResult = AppleScriptTask(APPLE_SCRIPT_FILE, "DownloadFile", templatePath & APPLE_SCRIPT_SPLIT_KEY & REPORT_TEMPLATE_URL)
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print IIf(Err.Number = 0, "    Download successful.", "    Error: " & Err.Description)
        #End If
        
        If downloadResult Then downloadResult = RequestFileAndFolderAccess(resourcesFolder, templatePath)
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
        Debug.Print "Attempting to load the SpeakingEvaluationTemplate.pptx."
    #End If
    
    templatePath = resourcesFolder & Application.PathSeparator & REPORT_TEMPLATE
    tempTemplatePath = GetTempFilePath(REPORT_TEMPLATE)
    
    DeleteFile tempTemplatePath ' Removing existing file to avoid problems overwriting

    If Not VerifyTemplateHash(templatePath) Then
        If Not DownloadReportTemplate(templatePath, resourcesFolder) Then
            msgToDisplay = "No template was found. Process canceled."
            msgResult = DisplayMessage(msgToDisplay, vbOKOnly, "Template Not Found", 150)
            LocateTemplate = ""
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print "    Unable to locate a copy of the template."
            #End If
            Exit Function
        End If
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

Private Function InstallFonts() As Boolean
    Const FONT_NAME As String = "Autumn in September.ttf"
    Const FONT_URL = "https://raw.githubusercontent.com/papercutter0324/SpeakingEvals/main/font.ttf"
    
    #If Mac Then
        InstallFonts = AppleScriptTask(APPLE_SCRIPT_FILE, "InstallFonts", FONT_NAME & APPLE_SCRIPT_SPLIT_KEY & FONT_URL)
    #Else
        Dim fso As Object, fontPath As String, sysFontPath As String
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        fontPath = Environ("LOCALAPPDATA") & "\Microsoft\Windows\Fonts\" & FONT_NAME
        sysFontPath = Environ("WINDIR") & "\Fonts\" & FONT_NAME
        
        If fso.FileExists(fontPath) Or fso.FileExists(sysFontPath) Then
            InstallFonts = True
            Exit Function
        End If
        
        Select Case True
            Case CheckForCurl()
                InstallFonts = DownloadUsingCurl(fontPath, FONT_URL)
            Case CheckForCurl()
                InstallFonts = DownloadUsingDotNet35(fontPath, FONT_URL)
            Case Else
                InstallFonts = False
        End Select
    #End If
End Function

Private Function VerifyTemplateHash(ByVal templatePath As String) As Boolean
    Const TEMPLATE_HASH As String = "97ca281d7fca39beb3f07555e6acde26"
    
    #If Mac Then
        VerifyTemplateHash = AppleScriptTask(APPLE_SCRIPT_FILE, "CompareMD5Hashes", templatePath & APPLE_SCRIPT_SPLIT_KEY & TEMPLATE_HASH)
    #Else
        Dim objShell As Object, shellOutput As String
        If Dir(templatePath) <> "" Then
            On Error GoTo ErrorHandler
            Set objShell = CreateObject("WScript.Shell")
            shellOutput = objShell.Exec("cmd /c certutil -hashfile """ & templatePath & """ MD5").StdOut.ReadAll
            VerifyTemplateHash = (LCase(TEMPLATE_HASH) = LCase(Trim$(Split(shellOutput, vbCrLf)(1))))
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

Private Function IsPptTemplateAlreadyOpen(ByVal resourcesFolder As String, ByVal REPORT_TEMPLATE As String) As Boolean
    Dim pptApp As Object, pptDoc As Object
    Dim templatePath As String, templateIsOpen As Boolean
    Dim pathOfOpenDoc As String
    Dim msgToDisplay As String
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Checking if a copy of the Speaking Evaluations Report template is already open."
    #End If
    
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    Err.Clear
    
    If Not pptApp Is Nothing Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Found an open instance of MS PowerPoint." & vbNewLine & _
                        "    Checking if template is open."
        #End If
        
        templatePath = resourcesFolder & Application.PathSeparator & REPORT_TEMPLATE
        
        For Each pptDoc In pptApp.Presentations
            pathOfOpenDoc = pptDoc.FullName
            ConvertOneDriveToLocalPath pathOfOpenDoc
            If StrComp(pathOfOpenDoc, templatePath, vbTextCompare) = 0 Then
                templateIsOpen = True
                #If PRINT_DEBUG_MESSAGES Then
                    Debug.Print "    Open instance of the template found. Asking if user wishes to automatically close and continue."
                #End If
                 msgToDisplay = "An open instance of MS PowerPoint has been detected. Please save any open files before continuing." & vbNewLine & vbNewLine & _
                                "Click OK to automatically close PowerPoint and continue, or click Cancel to finish and save your work."
                If DisplayMessage(msgToDisplay, vbOKCancel + vbCritical, "Error Loading PowerPoint", 310) = vbOK Then
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
        Debug.Print "    Open instance: " & templateIsOpen
    #End If
    
    Set pptDoc = Nothing
    Set pptApp = Nothing
    IsPptTemplateAlreadyOpen = templateIsOpen
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

Private Function SetSaveLocation(ByRef ws As Object, ByVal saveRoutine As String, ByVal resourcesFolder As String) As String
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
        permissionGranted = RequestFileAndFolderAccess(resourcesFolder, filePath)
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

Private Function FindPathToArchiveTool(ByVal resourcesFolder As String, Optional ByRef archiverName As String = "") As String
    Dim i As Integer, downloadResult As Boolean
    
    ' Declare OS specific variables, constants, and arrays
    #If Mac Then
        Dim scriptResultBoolean As Boolean
        
        Const RESOURCES_7ZIP_FILENAME As String = "7zz"
        Const RESOURCES_7ZIP_ARCHIVER_NAME As String = "Local 7zip"
    #Else
        Dim wshShell As Object
        Dim defaultPaths As Variant, archiverNames As Variant, exeNames As Variant, regKeys As Variant
        Dim regValue As String
    
        Const REG_KEY_7ZIP As String = "HKEY_LOCAL_MACHINE\SOFTWARE\7-Zip\Path"
        Const REG_KEY_7ZIP_32BIT As String = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\7-Zip\Path"
        
        Const ARCHIVER_NAME_7ZIP As String = "7Zip"
        
        Const EXE_NAME_7ZIP As String = "7z.exe"
        
        Const DEFAULT_PATH_7ZIP As String = "C:\Program Files\7-Zip\"
        Const DEFAULT_PATH_7ZIP_32Bit As String = "C:\Program Files (x86)\7-Zip\"
        
        Const RESOURCES_7ZIP_FILENAME As String = "7za.exe"
        Const RESOURCES_7ZIP_ARCHIVER_NAME As String = "Local 7zip"
        
        defaultPaths = Array(DEFAULT_PATH_7ZIP, DEFAULT_PATH_7ZIP_32Bit)
        archiverNames = Array(ARCHIVER_NAME_7ZIP, ARCHIVER_NAME_7ZIP)
        exeNames = Array(EXE_NAME_7ZIP, EXE_NAME_7ZIP)
        regKeys = Array(REG_KEY_7ZIP, REG_KEY_7ZIP_32BIT)
    #End If
    
    ' Find available archive utility
    #If Mac Then
        scriptResultBoolean = AppleScriptTask(APPLE_SCRIPT_FILE, "DoesFileExist", resourcesFolder & Application.PathSeparator & RESOURCES_7ZIP_FILENAME)
        If scriptResultBoolean Then
            scriptResultBoolean = AppleScriptTask(APPLE_SCRIPT_FILE, "ChangeFilePermissions", "+x" & APPLE_SCRIPT_SPLIT_KEY & resourcesFolder & Application.PathSeparator & RESOURCES_7ZIP_FILENAME)
            If scriptResultBoolean Then FindPathToArchiveTool = resourcesFolder & Application.PathSeparator & RESOURCES_7ZIP_FILENAME
            Exit Function
        End If
    #Else
        Set wshShell = CreateObject("WScript.Shell")
        
        ' First check default installation locations
        For i = LBound(defaultPaths) To UBound(defaultPaths)
            If Dir(defaultPaths(i) & exeNames(i)) <> "" Then
                archiverName = archiverNames(i)
                FindPathToArchiveTool = defaultPaths(i) & exeNames(i)
                Exit Function
            End If
        Next i
        
        ' If not found, check the registry for paths to custom locations
        On Error Resume Next
        For i = LBound(regKeys) To UBound(regKeys)
            regValue = wshShell.RegRead(regKeys(i))
            If Err.Number = 0 And regValue <> "" Then
                If Right(regValue, 1) <> "\" Then regValue = regValue & "\"
                
                ' Verify executable exists before returning path
                If Dir(regValue & exeNames(i)) <> "" Then
                    archiverName = archiverNames(i)
                    FindPathToArchiveTool = regValue & exeNames(i)
                    Exit Function
                End If
            End If
            Err.Clear
        Next i
        On Error GoTo 0
    #End If
    
    Download7Zip resourcesFolder, downloadResult
    
    If downloadResult Then
        archiverName = RESOURCES_7ZIP_ARCHIVER_NAME
        FindPathToArchiveTool = resourcesFolder & Application.PathSeparator & RESOURCES_7ZIP_FILENAME
    Else
        FindPathToArchiveTool = ""
    End If
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
    If libraryScriptsPath <> "" Then scriptResult = RequestFileAndFolderAccess(resourcesFolder, libraryScriptsPath)
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

Private Function RequestFileAndFolderAccess(ByVal resourcesFolder As String, Optional ByVal filePath As Variant = "") As Boolean
    Dim workingFolder As Variant, excelTempFolder As Variant, powerpointTempFolder As Variant
    Dim filePermissionCandidates As Variant, pathToRequest As Variant
    Dim fileAccessGranted As Boolean, allAccessHasBeenGranted As Boolean
    Dim i As Integer

    Select Case filePath
        Case ""
            workingFolder = ThisWorkbook.Path
            ConvertOneDriveToLocalPath workingFolder
            excelTempFolder = Environ("TMPDIR")
            powerpointTempFolder = Replace(excelTempFolder, "Excel", "PowerPoint")
            filePermissionCandidates = Array(workingFolder, resourcesFolder, excelTempFolder, powerpointTempFolder)
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
    Dim ws As Worksheet
    Dim enhancedDialogsStatus As String
    Dim enhancedDialogsAreDisabled As Boolean
    
    Set ws = ThisWorkbook.Sheets("MacOS Users")

    With ws.Shapes
        .Item("Button_SpeakingEvalsScpt_Missing").Visible = Not isAppleScriptInstalled
        .Item("Button_SpeakingEvalsScpt_Installed").Visible = isAppleScriptInstalled
        .Item("Button_DialogToolkit_Missing").Visible = Not isDialogToolkitInstalled
        .Item("Button_DialogToolkit_Installed").Visible = isDialogToolkitInstalled
        
        enhancedDialogsStatus = ws.Cells(1, 1).Value
        enhancedDialogsAreDisabled = Not isDialogToolkitInstalled Or enhancedDialogsStatus = "Enhanced Dialogs: Disabled" Or enhancedDialogsStatus = ""
        
        .Item("Button_EnhancedDialogs_Disable").Visible = enhancedDialogsAreDisabled
        .Item("Button_EnhancedDialogs_Enable").Visible = Not enhancedDialogsAreDisabled
    End With
    
    If enhancedDialogsAreDisabled And enhancedDialogsStatus <> "Enhanced Dialogs: Disabled" Then
        With ws
            .Unprotect
            .Cells(1, 1).Value = "Enhanced Dialogs: Disabled"
            .Protect
            .EnableSelection = xlUnlockedCells
        End With
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
    
    With ws
        .Unprotect
        .Cells(1, 1).Value = IIf(.Shapes("Button_EnhancedDialogs_Enable").Visible, SCRIPT_ENABLED, SCRIPT_DISABLED)
        .Protect
        .EnableSelection = xlUnlockedCells
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "    Value: """ & .Cells(1, 1).Value & """"
        #End If
    End With
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

Private Function DownloadUsingCurl(ByVal destinationPath As String, ByVal downloadURL As String) As Boolean
    Dim objShell As Object, fso As Object
    Dim downloadCommand As String
    
    On Error Resume Next
    Set objShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    downloadCommand = "cmd /c curl.exe -o """ & destinationPath & """ """ & downloadURL & """"
    objShell.Run downloadCommand, 0, True
    DownloadUsingCurl = fso.FileExists(destinationPath)
    
    #If PRINT_DEBUG_MESSAGES Then
        If Not DownloadUsingCurl Then Debug.Print "    curl download failed for " & downloadURL
    #End If
    On Error GoTo 0
End Function

Private Function DownloadUsingDotNet35(ByVal destinationPath As String, ByVal downloadURL As String) As Boolean
    Dim xmlHTTP As Object, fileStream As Object
    
    On Error Resume Next
    Set xmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Set fileStream = CreateObject("ADODB.Stream")
    
    xmlHTTP.Open "Get", downloadURL, False
    xmlHTTP.Send
    
    If xmlHTTP.Status = 200 Then
        fileStream.Open
        fileStream.Type = 1 ' Binary
        fileStream.Write xmlHTTP.responseBody
        fileStream.SaveToFile destinationPath, 2 ' Overwrite existing, if somehow present
        fileStream.Close
        DownloadUsingDotNet35 = True
    Else
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print "HTTP request failed. Status: " & xmlHTTP.Status & " - " & xmlHTTP.StatusText
        #End If
        DownloadUsingDotNet35 = False
    End If
    On Error GoTo 0
End Function
#End If
