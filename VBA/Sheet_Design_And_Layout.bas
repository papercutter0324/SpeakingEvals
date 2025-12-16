Option Explicit

#Const Windows = (Mac = 0)

Public Type LayoutData
    Name    As String
    Height  As Double
    Width   As Double
    Top     As Double
    Left    As Double
End Type

Public Type ButtonData
    Name        As String
    Height      As Double
    Width       As Double
    Top         As Double
    Left        As Double
    LeftBase    As Double
    Padding     As Double
    Spacing     As Double
End Type

Public Function SheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Sheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

Public Sub VerifySheetNames()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Sheets
        With ws
            Select Case .CodeName
                Case "Instructions"
                    .Name = "Instructions"
                Case "MacOS_Users"
                    .Name = "MacOS Users"
                Case "Options"
                    .Name = "Options"
                Case "Class_"
                    .Name = "Class"
            End Select
        End With
    Next ws
End Sub

Public Function SetLayoutInstructions() As Boolean
    Const BUTTON_HEIGHT             As Double = 70
    Const BUTTON_WIDTH              As Double = 200
    
    Dim instructionBoxes()   As LayoutData
    Dim instructionButtons   As ButtonData
    Dim buttonNames          As Variant
    Dim i                    As Long
    
    ToggleSheetProtection Instructions, False
    
    ' ReDim instructionBoxes(1 To 10)
    instructionBoxes() = PrepareInstructionTextboxData() 'UBound(instructionBoxes))
    
    For i = 1 To 10
        SetSheetTextBoxesLayout Instructions, instructionBoxes(i), True
    Next i
    
    buttonNames = Array("Button_Speadsheet", "Button_Font", "Button_ReportTemplate", "Button_SignatureTemplate", "Button_SourceCode")
    
    With Instructions.Shapes("Message - Seeing the Code")
        instructionButtons.Height = BUTTON_HEIGHT
        instructionButtons.Width = BUTTON_WIDTH
        instructionButtons.Top = .Top + .Height - instructionButtons.Height - 20
        instructionButtons.LeftBase = .Left + 20
        instructionButtons.Padding = BUTTON_WIDTH + (.Width - 40 - (BUTTON_WIDTH * 5)) / 4
    End With
    
    SetSheetButtonsLayout Instructions, buttonNames, instructionButtons, True
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.CodeExecution.Result", INDENT_LEVEL_2, IIf(Err.Number = 0, "Complete", "Errors found."))
    End If
    
    ToggleSheetProtection Instructions, True
    
    ' Figure out the desire error reporting method for this function
    ' Or, if most errors can simple be ignored, leave as is.
    SetLayoutInstructions = True
End Function

Public Function SetLayoutMacOSUsers() As Boolean
    Const BUTTON_HEIGHT As Double = 70
    Const BUTTON_WIDTH As Double = 200
    
    Dim shp As Shapes
    Dim enabledButtonNames As Variant
    Dim disabledButtonNames As Variant
    Dim i As Long
    
    Dim macOSBoxes() As LayoutData
    Dim macOSButtons As ButtonData
    
    ' ReDim macOSBoxes(1 To 3)
    macOSBoxes() = PrepareMacOSBoxes() 'UBound(macOSBoxes))
    
    For i = 1 To 3
        SetSheetTextBoxesLayout MacOS_Users, macOSBoxes(i), True
    Next i
    
    enabledButtonNames = Array("Button_SpeakingEvalsScpt_Installed", "Button_DialogToolkit_Installed", "Button_EnhancedDialogs_Enable")
    disabledButtonNames = Array("Button_SpeakingEvalsScpt_Missing", "Button_DialogToolkit_Missing", "Button_EnhancedDialogs_Disable")
    
    With MacOS_Users.Shapes("cURL_Command")
        macOSButtons.Height = BUTTON_HEIGHT
        macOSButtons.Width = BUTTON_WIDTH
        macOSButtons.Top = .Top + .Height + 8
        macOSButtons.LeftBase = .Left + 15
        macOSButtons.Padding = BUTTON_WIDTH + 15
    End With
    
    SetSheetButtonsLayout MacOS_Users, enabledButtonNames, macOSButtons
    SetSheetButtonsLayout MacOS_Users, disabledButtonNames, macOSButtons
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.CodeExecution.Result", INDENT_LEVEL_2, IIf(Err.Number = 0, "Complete", "Errors found."))
    End If

    SetLayoutMacOSUsers = True
End Function

Public Function SetLayoutOptions() As Boolean
    Const MAX_SIG_HEIGHT    As Double = 68
    Const MAX_SIG_WIDTH     As Double = 286

    Dim shp                       As Shape
    Dim shpName                   As String
    Dim sigShapeName              As String
    Dim sigEnabledShapeExists     As Boolean
    Dim sigDisabledShapeExists    As Boolean
    Dim sigPlaceholderShapeExists As Boolean
    Dim forceShapeVisibility      As Boolean
    Dim i                         As Long
    
    Dim optionsBoxes()        As LayoutData
    Dim optionsCertPreview    As LayoutData
    Dim optionsSigPlaceholder As LayoutData
    
    With Options.Columns
        .Item(1).ColumnWidth = 2
        .Item(2).ColumnWidth = 24
        .Item(3).ColumnWidth = 24
        .Item(4).ColumnWidth = 24
        .Item(5).ColumnWidth = 24
        .Item(6).ColumnWidth = 24
        .Item(7).ColumnWidth = 24
        .Item(12).ColumnWidth = 2
#If Mac Then
        .Item(8).ColumnWidth = 18
        .Item(9).ColumnWidth = 8
        .Item(10).ColumnWidth = 23
        .Item(11).ColumnWidth = 22
#Else
        .Item(8).ColumnWidth = 36.56
        .Item(9).ColumnWidth = 11.67
        .Item(10).ColumnWidth = 24
        .Item(11).ColumnWidth = 24
#End If
    End With
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.Worksheet.ValidateAllShapesArePresent")
    End If
    
    ReDim optionsBoxes(1 To 14)
    optionsBoxes() = PrepareOptionsTextBoxesData(MAX_SIG_HEIGHT, MAX_SIG_WIDTH)
    
    For i = 1 To 13
        If Not DoesShapeExist(Options, optionsBoxes(i).Name) And optionsBoxes(i).Name <> "mySignature-Placeholder" Then
            SetLayoutOptions = False
            Exit Function
        End If
        
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.Worksheet.ValidateShapePlacement", optionsBoxes(i).Name)
        End If
        
        forceShapeVisibility = (optionsBoxes(i).Name <> "Button_SignatureEmbedded" And optionsBoxes(i).Name <> "Button_SignatureMissing")
        
        SetSheetTextBoxesLayout Options, optionsBoxes(i), forceShapeVisibility
    Next i

    For Each shp In Options.Shapes
        shpName = shp.Name
        Select Case Split(shpName, "_")(0)
            Case "Embedded", "Layout"
                optionsBoxes(14).Name = shpName
                SetSheetTextBoxesLayout Options, optionsBoxes(14)
            Case "mySignature-Enabled", "mySignature-Disabled"
                ' Follow logic; is sigDisabledShapeExists needed?
                sigEnabledShapeExists = True
                sigShapeName = shpName
            Case "mySignature-Placeholder"
                sigPlaceholderShapeExists = True
                sigShapeName = shpName
        End Select
    Next shp
    
    Select Case True
        Case (sigEnabledShapeExists And sigPlaceholderShapeExists), (sigDisabledShapeExists And sigPlaceholderShapeExists)
            ' Throw an error
            ' Option to press Ok to automatically delete the placeholder
        Case sigEnabledShapeExists, sigDisabledShapeExists
            If g_UserOptions.EnableLogging Then
                DebugAndLogging GetMsg("Debug.Worksheet.ValidateShapeDimensions", sigShapeName)
            End If
            
            CenterAndFitSignature Options.Shapes(sigShapeName), optionsBoxes(4), MAX_SIG_HEIGHT, MAX_SIG_WIDTH
        Case sigPlaceholderShapeExists
            If g_UserOptions.EnableLogging Then
                DebugAndLogging GetMsg("Debug.Worksheet.ValidateShapeDimensions", sigShapeName)
            End If
            
            SetSheetTextBoxesLayout Options, optionsBoxes(13), True
        Case Else
            ' Throw an error
            ' Create a new placeholder shape
    End Select
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.CodeExecution.Result", vbNullString, IIf(Err.Number = 0, "Complete", "Errors found."))
    End If

    SetLayoutOptions = (Err.Number = 0)
End Function

Private Function PrepareInstructionTextboxData() As LayoutData()
    Const PADDING_LEFT              As Double = 15
    Const PADDING_TOP               As Double = 15
    Const SHAPE_SPACING             As Double = 20
    
    Const WIDE_WIDTH                As Double = 1525
    Const MEDIUM_WIDTH              As Double = 1165
    Const THIN_WIDTH                As Double = 340
    
    Const TITLEBAR_HEIGHT           As Double = 58
    Const INSTRUCTIONS_MSG_HEIGHT   As Double = 560
    Const CODE_MSG_HEIGHT           As Double = 265
    Const WARNING_MSG_HEIGHT        As Double = 150
    Const IMPORTANT_MSG_HEIGHT      As Double = 675
    Const TODO_MSG_HEIGHT           As Double = 585

    Dim tmpData() As LayoutData
    ReDim tmpData(1 To 10) ' numberOfElements)
    
    tmpData(1) = SetLayoutValues("Title Bar - Instructions", _
                                 TITLEBAR_HEIGHT, _
                                 MEDIUM_WIDTH, _
                                 PADDING_TOP, _
                                 PADDING_LEFT)
    
    tmpData(2) = SetLayoutValues("Message - Instructions", _
                                 INSTRUCTIONS_MSG_HEIGHT, _
                                 MEDIUM_WIDTH, _
                                 tmpData(1).Top + TITLEBAR_HEIGHT, _
                                 tmpData(1).Left)
                    
    tmpData(3) = SetLayoutValues("Title Bar - Seeing the Code", _
                                 TITLEBAR_HEIGHT, _
                                 MEDIUM_WIDTH, _
                                 tmpData(2).Top + tmpData(2).Height + SHAPE_SPACING, _
                                 tmpData(1).Left)
    
    tmpData(4) = SetLayoutValues("Message - Seeing the Code", _
                                 CODE_MSG_HEIGHT, _
                                 MEDIUM_WIDTH, _
                                 tmpData(3).Top + tmpData(3).Height, _
                                 tmpData(3).Left)
                    
    tmpData(5) = SetLayoutValues("Title Bar - Warning", _
                                 TITLEBAR_HEIGHT, _
                                 THIN_WIDTH, _
                                 tmpData(1).Top, _
                                 tmpData(1).Left + tmpData(1).Width + SHAPE_SPACING)
                    
    tmpData(6) = SetLayoutValues("Message - Warning", _
                                 WARNING_MSG_HEIGHT, _
                                 THIN_WIDTH, _
                                 tmpData(5).Top + tmpData(5).Height, _
                                 tmpData(5).Left)
                    
    tmpData(7) = SetLayoutValues("Title Bar - Important Files", _
                                 TITLEBAR_HEIGHT, _
                                 THIN_WIDTH, _
                                 tmpData(6).Top + tmpData(6).Height + SHAPE_SPACING, _
                                 tmpData(5).Left)
                    
    tmpData(8) = SetLayoutValues("Message - Important Files", _
                                 IMPORTANT_MSG_HEIGHT, _
                                 THIN_WIDTH, _
                                 tmpData(7).Top + tmpData(7).Height, _
                                 tmpData(7).Left)
                    
    tmpData(9) = SetLayoutValues("Title Bar - ToDo", _
                                 TITLEBAR_HEIGHT, _
                                 WIDE_WIDTH, _
                                 tmpData(4).Top + tmpData(4).Height + SHAPE_SPACING, _
                                 tmpData(1).Left)
                    
    tmpData(10) = SetLayoutValues("Message - ToDo", _
                                  TODO_MSG_HEIGHT, _
                                  WIDE_WIDTH, _
                                  tmpData(9).Top + tmpData(9).Height, _
                                  tmpData(9).Left)
    PrepareInstructionTextboxData = tmpData()
End Function

Private Function PrepareMacOSBoxes() As LayoutData()
    Const TITLEBAR_HEIGHT   As Double = 58
    Const MACOS_MSG_HEIGHT  As Double = 700
    Const CURL_HEIGHT       As Double = 60
    Const WIDE_WIDTH        As Double = 1300
    Const CURL_WIDTH        As Double = 660
    Const PADDING_TOP       As Double = 15
    Const PADDING_LEFT      As Double = 15

    Dim tmpData() As LayoutData
    ReDim tempData(1 To 3) 'numberOfElements)

    tmpData(1) = SetLayoutValues("Title Bar", _
                                 TITLEBAR_HEIGHT, _
                                 WIDE_WIDTH, _
                                 PADDING_TOP, _
                                 PADDING_LEFT)
    
    tmpData(2) = SetLayoutValues("Message", _
                                 MACOS_MSG_HEIGHT, _
                                 tmpData(1).Width, _
                                 tmpData(1).Top + tmpData(1).Height, _
                                 tmpData(1).Left)
    
    tmpData(3) = SetLayoutValues("cURL_Command", _
                                 CURL_HEIGHT, _
                                 CURL_WIDTH, _
                                 tmpData(2).Top, _
                                 tmpData(1).Left + tmpData(1).Width - tmpData(3).Width)

    PrepareMacOSBoxes = tmpData()
End Function

Private Function PrepareOptionsTextBoxesData(ByVal maxSigHeight As Double, ByVal maxSigWidth As Double) As LayoutData()
    Const TITLEBAR_HEIGHT       As Double = 58
    Const SIG_CONTAINER_HEIGHT  As Double = 86
    Const CERTS_MSG_HEIGHT      As Double = 340
    Const MSG_TB_WIDTH          As Double = 690
    Const SIG_TB_WIDTH          As Double = 300
    Const CERTS_TB_WIDTH        As Double = 620
    Const CERT_PREVIEW_HEIGHT   As Double = 300
    Const CERT_PREVIEW_WIDTH    As Double = 433.32
    Const BTN_CONTAINER_WIDTH   As Double = 205
    Const PADDING_TOP           As Double = 15
    Const PADDING_LEFT          As Double = 15
    Const MSG_HEIGHT            As Double = 400
    Const BTN_SIGNATURE_HEIGHT  As Double = 65
    Const BTN_SIGNATURE_WIDTH   As Double = 175
    Const TB_ADV_SETTINGS_WIDTH As Double = 345
    
    Dim tmpData() As LayoutData
    ReDim tmpData(1 To 14) ' numberOfElements)

    tmpData(1) = SetLayoutValues("Title Bar", _
                                 TITLEBAR_HEIGHT, _
                                 MSG_TB_WIDTH, _
                                 PADDING_TOP, _
                                 PADDING_LEFT)
                                          
    tmpData(2) = SetLayoutValues("Message", _
                                 MSG_HEIGHT, _
                                 tmpData(1).Width + SIG_TB_WIDTH, _
                                 tmpData(1).Top + tmpData(1).Height, _
                                 tmpData(1).Left)

    tmpData(3) = SetLayoutValues("Signature Title Bar", _
                                 tmpData(1).Height, _
                                 SIG_TB_WIDTH, _
                                 tmpData(1).Top, _
                                 tmpData(1).Left + tmpData(1).Width)
                                          
    tmpData(4) = SetLayoutValues("Signature Container", _
                                 SIG_CONTAINER_HEIGHT, _
                                 tmpData(3).Width, _
                                 tmpData(3).Top + tmpData(3).Height, _
                                 tmpData(3).Left)
    
    tmpData(5) = SetLayoutValues("Button_Container", _
                                 tmpData(4).Height, _
                                 BTN_CONTAINER_WIDTH, _
                                 tmpData(4).Top + tmpData(4).Height, _
                                 tmpData(4).Left + (tmpData(4).Width / 2) - (BTN_CONTAINER_WIDTH / 2))
                                       
    tmpData(6) = SetLayoutValues("Button_SignatureEmbedded", _
                                 BTN_SIGNATURE_HEIGHT, _
                                 BTN_SIGNATURE_WIDTH, _
                                 tmpData(5).Top + tmpData(5).Height - (tmpData(5).Height / 2) - (BTN_SIGNATURE_HEIGHT / 2), _
                                 tmpData(5).Left + tmpData(5).Width - (tmpData(5).Width / 2) - (BTN_SIGNATURE_WIDTH / 2))
    
    tmpData(7) = SetLayoutValues("Button_SignatureMissing", _
                                 tmpData(6).Height, _
                                 tmpData(6).Width, _
                                 tmpData(6).Top, _
                                 tmpData(6).Left)
    
    tmpData(8) = SetLayoutValues("Advanced_TitleBar", _
                                 tmpData(1).Height, _
                                 TB_ADV_SETTINGS_WIDTH, _
                                 tmpData(1).Top, _
                                 tmpData(3).Left + tmpData(3).Width + 10)
                                       
    tmpData(9) = SetLayoutValues("Advanced_Message", _
                                 tmpData(2).Height, _
                                 TB_ADV_SETTINGS_WIDTH, _
                                 tmpData(8).Top + tmpData(8).Height, _
                                 tmpData(8).Left)

    tmpData(10) = SetLayoutValues("Certificate_TitleBar", _
                                  tmpData(1).Height, _
                                  CERTS_TB_WIDTH, _
                                  tmpData(2).Top + tmpData(2).Height + PADDING_TOP, _
                                  tmpData(2).Left)
    
    tmpData(11) = SetLayoutValues("Certificate_Message", _
                                  CERTS_MSG_HEIGHT, _
                                  tmpData(10).Width, _
                                  tmpData(10).Top + tmpData(10).Height, _
                                  tmpData(10).Left)
                                       
    tmpData(12) = SetLayoutValues("Certificate_Options_TitleBar", _
                                  tmpData(10).Height, _
                                  tmpData(8).Left + tmpData(8).Width - tmpData(1).Left - tmpData(10).Width, _
                                  tmpData(10).Top, _
                                  tmpData(10).Left + tmpData(10).Width)
                                       
    tmpData(13) = SetLayoutValues("mySignature-Placeholder", _
                                  maxSigHeight, _
                                  maxSigWidth, _
                                  tmpData(4).Top + (tmpData(4).Height - maxSigHeight) / 2, _
                                  tmpData(4).Left + (tmpData(4).Width - maxSigWidth) / 2)
    
    tmpData(14) = SetLayoutValues(vbNullString, _
                                  CERT_PREVIEW_HEIGHT, _
                                  CERT_PREVIEW_WIDTH, _
                                  tmpData(11).Top + (tmpData(11).Height / 2) - (CERT_PREVIEW_HEIGHT / 2), _
                                  tmpData(12).Left + 20)

    PrepareOptionsTextBoxesData = tmpData()
End Function

Private Function SetLayoutValues(ByVal shapeName As String, ByVal heightValue As Double, ByVal widthValue As Double, ByVal topValue As Double, ByVal leftValue As Double) As LayoutData
    Dim tmpData As LayoutData
    
    With tmpData
        .Name = shapeName
        .Height = heightValue
        .Width = widthValue
        .Top = topValue
        .Left = leftValue
    End With

    SetLayoutValues = tmpData
End Function

Private Sub SetSheetTextBoxesLayout(ByRef ws As Worksheet, ByRef targetShape As LayoutData, Optional ByVal forceVisibility As Boolean = False)
    Dim shp As Shape
    
    On Error GoTo MissingShape
    Set shp = ws.Shapes(targetShape.Name)
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        With shp
            If forceVisibility Then .Visible = msoTrue
            
            .Top = targetShape.Top
            .Left = targetShape.Left
            
            .LockAspectRatio = msoFalse
            .Height = targetShape.Height
            .Width = targetShape.Width
            .LockAspectRatio = msoTrue
        End With
    End If
MissingShape:
End Sub

Private Sub SetSheetButtonsLayout(ByRef ws As Worksheet, ByRef buttonNames As Variant, ByRef sheetButtonData As ButtonData, Optional ByVal forceVisibility As Boolean = False)
    Dim shp As Shape
    Dim i As Long
    
    For i = LBound(buttonNames) To UBound(buttonNames)
        On Error GoTo MissingShape
        Set shp = ws.Shapes(buttonNames(i))
        On Error GoTo 0
        
        If Not shp Is Nothing Then
            With shp
                If forceVisibility Then .Visible = msoTrue
                
                .Top = sheetButtonData.Top
                .Left = sheetButtonData.LeftBase + (sheetButtonData.Padding * i)
                
                .LockAspectRatio = msoFalse
                .Height = sheetButtonData.Height
                .Width = sheetButtonData.Width
                .LockAspectRatio = msoTrue
            End With
        End If
    Next i
MissingShape:
End Sub

Public Sub SetLayoutClassRecords(ByRef ws As Worksheet) ' As Boolean
    Const VISIBLE_BORDERS_RANGES    As String = "A1:C6,D1:I6,J1:J6,K1:M6,A7:A7,B7:C7,D7:I7,J7:J7,K7:M7,A8:A32,B8:C32,D8:I32,J8:J32,K8:M32"
    Const HIDDEN_BORDERS_RANGES     As String = "D1:I6,J1:J6,K1:M6"
    Const LOCKED_CELLS_RANGES       As String = "A1:B6,A7:M7,D1:J6,K1:M1,K2:K6,L5:L6,M2:M6,A8:A32"
    Const UNLOCKED_CELLS_RANGES     As String = "C1:C6,B8:M32,L2:L4"
    
    Const CLASS_INFO_CELLS        As String = "A1:C6"
    Const NAMELIST_HEADER_CELLS   As String = "A7:C7"
    Const GRAMMAR_HEADER_CELLS    As String = "D1:D7"
    Const PRONUN_HEADER_CELLS     As String = "E1:E7"
    Const FLUENCY_HEADER_CELLS    As String = "F1:F7"
    Const MANNER_HEADER_CELLS     As String = "G1:G7"
    Const CONTENT_HEADER_CELLS    As String = "H1:H7"
    Const EFFORT_HEADER_CELLS     As String = "I1:I7"
    Const BUTTON_HEADER_CELLS     As String = "J1:J6"
    Const COMMENT_HEADER_CELLS    As String = "J7"
    Const WINNER_HEADER_CELLS     As String = "K1:M6" ' "K1,K2:M6,K7"
    Const NOTES_HEADER_CELLS      As String = "K7:M7"
    Const NAME_INDEX_CELLS        As String = "A8:A32"
    Const NAME_LISTS_CELLS        As String = "B8:C32"
    Const GRAMMAR_ENTRY_CELLS     As String = "D8:D32"
    Const PRONUN_ENTRY_CELLS      As String = "E8:E32"
    Const FLUENCY_ENTRY_CELLS     As String = "F8:F32"
    Const MANNER_ENTRY_CELLS      As String = "G8:G32"
    Const CONTENT_ENTRY_CELLS     As String = "H8:H32"
    Const EFFORT_ENTRY_CELLS      As String = "I8:I32"
    Const COMMENT_ENTRY_CELLS     As String = "J8:J32"
    Const NOTE_ENTRY_CELLS        As String = "K8:M32"
    
    Const SHADING_RANGE_NAME    As Long = 0
    Const SHADING_RANGE_ADDRESS As Long = 1
    Const SHADING_COLOR_VALUES  As Long = 2

    Dim visibleBorders  As Range
    Dim hiddenBorders   As Range
    Dim lockedCells     As Range
    Dim unlockedCells   As Range
    Dim existingName    As Name
    Dim shadingRanges   As Variant
    Dim i               As Long

    shadingRanges = Array( _
        Array("Class_Info", CLASS_INFO_CELLS, CellShading.White), _
        Array("Name_List_Header", NAMELIST_HEADER_CELLS, CellShading.LightSteal), _
        Array("Grammar_Header", GRAMMAR_HEADER_CELLS, CellShading.Purple), _
        Array("Pronunciation_Header", PRONUN_HEADER_CELLS, CellShading.Teal), _
        Array("Fluency_Header", FLUENCY_HEADER_CELLS, CellShading.DeepPink), _
        Array("Manner_Header", MANNER_HEADER_CELLS, CellShading.Orange), _
        Array("Content_Header", CONTENT_HEADER_CELLS, CellShading.Green), _
        Array("Effort_Header", EFFORT_HEADER_CELLS, CellShading.MediumBlue), _
        Array("Button_Header", BUTTON_HEADER_CELLS, CellShading.White), _
        Array("Comment_Header", COMMENT_HEADER_CELLS, CellShading.Grey), _
        Array("Winner_Header", WINNER_HEADER_CELLS, CellShading.Tan), _
        Array("Notes_Header", NOTES_HEADER_CELLS, CellShading.Tan), _
        Array("Name_List_Index", NAME_INDEX_CELLS, CellShading.mediumGrey), _
        Array("Name_List_Entry", NAME_LISTS_CELLS, CellShading.White), _
        Array("Grammar_Entry", GRAMMAR_ENTRY_CELLS, CellShading.Lavender), _
        Array("Pronunciation_Entry", PRONUN_ENTRY_CELLS, CellShading.SkyBlue), _
        Array("Fluency_Entry", FLUENCY_ENTRY_CELLS, CellShading.LightPink), _
        Array("Manner_Entry", MANNER_ENTRY_CELLS, CellShading.LightPeach), _
        Array("Content_Entry", CONTENT_ENTRY_CELLS, CellShading.LightGreen), _
        Array("Effort_Entry", EFFORT_ENTRY_CELLS, CellShading.LightBlue), _
        Array("Comment_Entry", COMMENT_ENTRY_CELLS, CellShading.LightGrey), _
        Array("Note_Entry", NOTE_ENTRY_CELLS, CellShading.Beige) _
    )

    ' Step 1: Apply Bulk Formatting & Named Ranges
    With ws
        Set visibleBorders = .Range(VISIBLE_BORDERS_RANGES)
        Set hiddenBorders = .Range(HIDDEN_BORDERS_RANGES)
        Set lockedCells = .Range(LOCKED_CELLS_RANGES)
        Set unlockedCells = .Range(UNLOCKED_CELLS_RANGES)
        
        With .Columns
            .Item("A").ColumnWidth = 7
            .Item("B:C").ColumnWidth = 18
            .Item("D:I").ColumnWidth = 22
            .Item("J").ColumnWidth = 105
            .Item("K").ColumnWidth = 20
            .Item("L").ColumnWidth = 56
            .Item("M").ColumnWidth = 2
        End With
        
        With .Rows
            .Item("1:6").RowHeight = 30
            .Item("7").RowHeight = 25
            .Item("8:32").RowHeight = 65
        End With
        
        For i = LBound(shadingRanges) To UBound(shadingRanges)
            .Range(shadingRanges(i)(SHADING_RANGE_ADDRESS)).Interior.Color = shadingRanges(i)(SHADING_COLOR_VALUES)
        Next i
    End With
    
    With visibleBorders.Borders
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Item(xlInsideHorizontal).LineStyle = xlDash
        .Item(xlInsideHorizontal).Weight = xlThin
        .Item(xlInsideVertical).LineStyle = xlLineStyleNone
    End With
    
    With hiddenBorders.Borders
        .Item(xlInsideHorizontal).LineStyle = xlLineStyleNone
        .Item(xlInsideVertical).LineStyle = xlLineStyleNone
    End With
    
    With unlockedCells
        .Locked = False
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .WrapText = True
        .NumberFormat = "@" ' Treat as text
    End With
    
    lockedCells.Locked = True
    
    ' Set Colors and Named Ranges
    With ThisWorkbook.Names
        For i = LBound(shadingRanges) To UBound(shadingRanges)
            Set existingName = Nothing
            On Error Resume Next
            Set existingName = .Item(shadingRanges(i)(SHADING_RANGE_NAME))
            If Not existingName Is Nothing Then existingName.Delete
            On Error GoTo 0

            .Add Name:=shadingRanges(i)(SHADING_RANGE_NAME), RefersTo:=ws.Range(shadingRanges(i)(SHADING_RANGE_ADDRESS))
        Next i
    End With

    SetLayoutClassRecordsButtons ws

    ' Separate Steps for clarity
    VerifyValidationAndFontsForClassRecords ws
    
    UpdateWinnersLists ws
    
    SetDefaultShading ws
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.CodeExecution.Result", INDENT_LEVEL_2, IIf(Err.Number = 0, "Complete", "Errors found."))
    End If

    ' Figure out the desire error reporting method for this function
    ' Or, if most errors can simple be ignored, leave as is.
    ' SetLayoutClassRecords = True
End Sub

Public Sub SetLayoutClassRecordsButtons(ByRef ws As Worksheet)
    Dim cellTop As Double
    Dim cellHeight As Double
    Dim cellLeft As Double
    Dim cellWidth As Double
    Dim cellVerticalSpacing As Double
    Dim cellHorizontalSpacing As Double
    Dim buttonProps As Variant
    Dim i As Long
    
    Const BUTTON_HEIGHT As Double = 45
    Const BUTTON_WIDTH As Double = 195
    
    With ws.Cells.Item(1, 10)
        cellTop = .Top
        cellHeight = .Height * 6
        cellLeft = .Left
        cellWidth = .Width
    End With
    
    cellVerticalSpacing = (cellHeight - (3 * BUTTON_HEIGHT)) / 5 + 2
    cellHorizontalSpacing = (cellWidth - (2 * BUTTON_WIDTH)) / 3
    
    ' Define button properties in an array: {Button Name, Row Index, Col Index}
    buttonProps = Array( _
        Array("Button_GenerateProofs", 1, 1), _
        Array("Button_GenerateReports", 1, 2), _
        Array("Button_AutoSelectWinners", 2, 1), _
        Array("Button_GenerateCertificates", 2, 2), _
        Array("Button_OpenTypingSite", 3, 1), _
        Array("Button_RepairLayout", 3, 2) _
    )

    ' Loop through button array and set positions
    With ws.Shapes
        For i = LBound(buttonProps) To UBound(buttonProps)
            SetShapeDimensionAndPosition .Item(buttonProps(i)(0)), _
                                         BUTTON_HEIGHT, _
                                         BUTTON_WIDTH, _
                                         cellTop + (cellVerticalSpacing * buttonProps(i)(1)) + ((buttonProps(i)(1) - 1) * BUTTON_HEIGHT), _
                                         cellLeft + (cellHorizontalSpacing * buttonProps(i)(2)) + ((buttonProps(i)(2) - 1) * BUTTON_WIDTH), _
                                         True
        Next i
    End With
End Sub

Public Sub RepairLayouts(ByRef ws As Worksheet)
    Dim gradeCell As Range
    Dim currentValue As String
    Dim newValue As String
    
    ToggleSheetProtection ws, False
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.Worksheet.RepairLayout", ws.Name)
    End If
    
    SetLayoutClassRecords ws
    
    For Each gradeCell In ws.Range(g_STUDENT_GRADES)
        currentValue = CStr(gradeCell.Value)
        newValue = FormatGrade(currentValue, True)
        WriteNewRangeValue gradeCell, newValue
    Next gradeCell
    
    ' validate winners list names?
        
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.Worksheet.RepairLayoutComplete")
    End If
    
    ToggleSheetProtection ws, True
End Sub

Public Function ValidateSheetLayoutsOnLoad() As Boolean
    Dim ws As Worksheet
    ' Dim errorEncountered As Boolean

    For Each ws In ThisWorkbook.Worksheets
        ToggleSheetProtection ws, False
        
        With ws
            If g_UserOptions.EnableLogging Then
                DebugAndLogging GetMsg("Debug.Worksheet.PrintWorksheetName", INDENT_LEVEL_1, .Name)
            End If
            
            Select Case .Name
                Case "Instructions", "Class"
                    ' Nothing to do...
                    ' Kinda hacky, but this can be made nicer later
                Case "MacOS Users"
                    #If Mac Then
                        On Error Resume Next
                        .Shapes("cURL_Command").TextFrame2.TextRange.Characters.text = GetMsg("Textbox.MacOS.CurlCommand.Text")
                        On Error GoTo 0
                        If Err.Number <> 0 Then
                            'Display and log error
                            Err.Clear
                        End If
                    #End If
                Case "Options"
                    SetLayoutOptions
                Case Else
                    AutoPopulateEvaluationDateValues ws ', errorEncountered
                    SetLayoutClassRecords ws ', errorEncountered
            End Select
        End With
        
        ToggleSheetProtection ws, True

        ' If errorEncountered Then Exit For
    Next ws

    ValidateSheetLayoutsOnLoad = True ' Not errorEncountered
End Function

Public Sub AutoPopulateEvaluationDateValues(ByRef ws As Worksheet)
    Dim dateCell As Range
    Dim dateCellValue As String
    ' Dim dateAsDate As Date
    ' Dim dateToCheck As Date
    
    On Error Resume Next
    With ws
        If .CodeName <> "Class_" And .Range("A6").Value = "Evaluation Date:" Then
            Set dateCell = .Range(g_EVALUATION_DATE)
            dateCellValue = dateCell.Value
    
            If Len(Trim$(dateCellValue)) = 0 Then
                WriteNewRangeValue dateCell, Format$(Date, GetLocaleDateOrder)
            ElseIf IsDate(Trim$(dateCellValue)) Then
                ' dateAsDate = CDate(Trim$(dateCell.Value))
                ' dateToCheck = DateAdd("m", -2, Date)
                
                ' If dateAsDate < dateToCheck Then
                '     dateCell.Value = Format$(Date, "MMM. YYYY")
                ' End If
            Else
                DisplayMessage "Display.Worksheet.InvalidDate", .Name
                WriteNewRangeValue dateCell, vbNullString
            End If
        End If
    End With
    On Error GoTo 0
End Sub

Public Function DoesShapeExist(ByRef ws As Worksheet, ByVal shapeName As String) As Boolean
    Dim shapeExists As Boolean
    
    On Error Resume Next
    shapeExists = Not ws.Shapes(shapeName) Is Nothing
    On Error GoTo 0
    
    If g_UserOptions.EnableLogging And Not shapeExists Then
        DebugAndLogging GetMsg("Debug.Worksheet.MissingShape", shapeName)
    End If
    
    DoesShapeExist = shapeExists
End Function

Public Sub OptionsShapeVisibility(ByRef ws As Worksheet)
    Dim shp As Shapes
    Dim defaultOption As String
    Dim signaturePresent As Boolean
    Dim certificateOptions As Variant
    Dim optionRangeRow As Long
    Dim i As Long
    
    Set shp = ws.Shapes
    
    ' Step 1: Verify correct signature button is displayed
    On Error Resume Next
    signaturePresent = DoesShapeExist(Options, "mySignature-Enabled")
    On Error GoTo 0
    
    shp("Button_SignatureEmbedded").Visible = signaturePresent
    shp("Button_SignatureMissing").Visible = Not signaturePresent
    
    ' Step 2: Check certificate options, setting defaults as needed, and generate preview
    certificateOptions = ws.Range("J10:K15")
    For i = LBound(certificateOptions, 1) To UBound(certificateOptions, 1)
        If certificateOptions(i, 2) = vbNullString Then
            optionRangeRow = i + 9
            defaultOption = GetDefaultCertificateOptions(ws.Range("J" & optionRangeRow).Value)
            WriteNewRangeValue ws.Range("k" & optionRangeRow), defaultOption
        End If
    Next i
End Sub

Public Sub SetShapeDimensionAndPosition(ByVal buttonShape As Shape, ByVal buttonHeight As Double, ByVal buttonWidth As Double, ByVal buttonTop As Double, ByVal buttonLeft As Double, Optional ByVal forceVisibility As Boolean = False)
    With buttonShape
        If forceVisibility And .Visible <> msoTrue Then .Visible = msoTrue
        .LockAspectRatio = msoFalse
        .Height = buttonHeight
        .Width = buttonWidth
        .LockAspectRatio = msoTrue
        .Top = buttonTop
        .Left = buttonLeft
    End With
End Sub

Private Sub CenterAndFitSignature(ByVal shp As Shape, ByRef sigContainer As LayoutData, ByVal maxHeight As Double, ByVal maxWidth As Double)
    Dim sigImage As LayoutData
    Dim aspectRatio As Double

    sigImage.Name = shp.Name
    sigImage.Height = maxHeight
    sigImage.Width = maxWidth
    aspectRatio = shp.Width / shp.Height
    
    If (maxWidth / maxHeight) > aspectRatio Then
        sigImage.Width = maxHeight * aspectRatio
    Else
        sigImage.Height = maxWidth / aspectRatio
    End If

    sigImage.Top = sigContainer.Top + (sigContainer.Height - sigImage.Height) / 2
    sigImage.Left = sigContainer.Left + (sigContainer.Width - sigImage.Width) / 2

    SetSheetTextBoxesLayout Options, sigImage, True
End Sub

Public Sub VerifyValidationAndFontsForClassRecords(ByRef ws As Worksheet)
    Const LEVEL_LIST          As String = "Theseus,Perseus,Odysseus,Hercules,Artemis,Hermes,Apollo,Zeus,E5 Athena,Helios,Poseidon,Gaia,Hera,E6 Song's"
    Const DAYS_LIST           As String = "MonWed,MonFri,WedFri,MWF,TTh,MWF (Class A),MWF (Class B),TTh (Class A),TTh (Class B)"
    Const TIME_LIST           As String = "9pm,830pm,8pm,7pm,6pm,530pm,5pm,4pm,3pm,2pm,1pm,12pm,11am,10am,9am"
    Const ENGLISH_FONT        As String = "Calibri"
    Const KOREAN_FONT_DEFAULT As String = "Malgun Gothic"
    Const KOREAN_FONT_CUSTOM  As String = "Kakao Big Sans"

    Const ENGLISH_TEACHER_INPUT_TITLE   As String = "Native Teacher's Name"
    Const ENGLISH_TEACHER_INPUT_MESSAGE As String = "Please enter just your" & vbNewLine & "name, no suffix or title" & vbNewLine & "like ""tr."""
    Const KOREAN_TEACHER_INPUT_TITLE    As String = "Korean Teacher's Name"
    Const KOREAN_TEACHER_INPUT_MESSAGE  As String = "Please write their Korean name. The parents are unlikely to know their English name."
    Const CLASS_LEVEL_INPUT_TITLE       As String = "Class Level"
    Const CLASS_LEVEL_INPUT_MESSAGE     As String = "Click on the down arrow and choose the class's level from the list." & vbNewLine & vbNewLine & "Scroll for more options."
    Const CLASS_DAYS_INPUT_TITLE        As String = "Class Days"
    Const CLASS_DAYS_INPUT_MESSAGE      As String = "Click on the down arrow and select the days when you see this class." & vbNewLine & vbNewLine & "'Class A' and 'Class B' are intended for Athena, Song's, and other split classes."
    Const CLASS_TIME_INPUT_TITLE        As String = "Class Time"
    Const CLASS_TIME_INPUT_MESSAGE      As String = "Click on the down arrow and select what time your first class is with them each week." & vbNewLine & vbNewLine & "Scroll for more options."
    Const DATE_INPUT_TITLE              As String = "Date Format"
    Const STUDENT_ENG_NAME_TITLE        As String = "Character Limit"
    Const STUDENT_ENG_NAME_MESSAGE      As String = "20 characters or fewer recommended."
    Const STUDENT_KOR_NAME_TITLE        As String = "Language Reminder"
    Const STUDENT_KOR_NAME_MESSAGE      As String = "Please write their names in Korean."
    Const STUDENT_GRADE_TITLE           As String = "Enter a Grade"
    Const STUDENT_GRADE_MESSAGE         As String = "Valid Letter Grades" & vbNewLine & "  A+ / A / B+ / B / C" & vbNewLine & vbNewLine & "Valid Numeric Scores   " & vbNewLine & "  1 ~ 5"
    Const COMMENTS_TITLE                As String = "Character Limit"
    Const COMMENTS_MESSAGE              As String = "960 characters"

    Dim cellRng             As Range
    Dim validationValues    As Variant
    Dim dateInputMessage    As String
    Dim availablekoreanFont As String
    Dim i                   As Long

    Dim engTeacherValidation    As ValidationSettings
    Dim korTeacherValidation    As ValidationSettings
    Dim classLevelValidation    As ValidationSettings
    Dim classDaysValidation     As ValidationSettings
    Dim classTimeValidation     As ValidationSettings
    Dim evalDateValidation      As ValidationSettings
    Dim englishNameValidation   As ValidationSettings
    Dim koreanNameValidation    As ValidationSettings
    Dim studentGradeValidation  As ValidationSettings
    Dim commentsValidation      As ValidationSettings
    Dim teacherNotesValidation  As ValidationSettings
    
    Select Case Application.International(xlDateOrder)
       Case 0: dateInputMessage = "MM/DD/YYYY"
       Case 1: dateInputMessage = "DD/MM/YYYY"
       Case 2: dateInputMessage = "YYYY/MM/DD"
    End Select

    availablekoreanFont = IIf(g_UserOptions.AllFontsAreInstalled, KOREAN_FONT_CUSTOM, KOREAN_FONT_DEFAULT)

    With engTeacherValidation
        .TypeOfValidation = xlValidateInputOnly
        .AlertStyle = xlValidAlertStop
        .InputTitle = ENGLISH_TEACHER_INPUT_TITLE
        .InputMessage = ENGLISH_TEACHER_INPUT_MESSAGE
        .ShowInput = g_UserOptions.DisplayEntryTips
        .ShowError = False
    End With

    With korTeacherValidation
        .TypeOfValidation = xlValidateInputOnly
        .AlertStyle = xlValidAlertStop
        .InputTitle = KOREAN_TEACHER_INPUT_TITLE
        .InputMessage = KOREAN_TEACHER_INPUT_MESSAGE
        .ShowInput = g_UserOptions.DisplayEntryTips
        .ShowError = False
    End With

    With classLevelValidation
        .TypeOfValidation = xlValidateList
        .AlertStyle = xlValidAlertStop
        .InputTitle = CLASS_LEVEL_INPUT_TITLE
        .InputMessage = CLASS_LEVEL_INPUT_MESSAGE
        .Formula1 = LEVEL_LIST
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = g_UserOptions.DisplayEntryTips
        .ShowError = False
    End With

    With classDaysValidation
        .TypeOfValidation = xlValidateList
        .AlertStyle = xlValidAlertStop
        .InputTitle = CLASS_DAYS_INPUT_TITLE
        .InputMessage = CLASS_DAYS_INPUT_MESSAGE
        .Formula1 = DAYS_LIST
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = g_UserOptions.DisplayEntryTips
        .ShowError = False
    End With

    With classTimeValidation
        .TypeOfValidation = xlValidateList
        .AlertStyle = xlValidAlertStop
        .InputTitle = CLASS_TIME_INPUT_TITLE
        .InputMessage = CLASS_TIME_INPUT_MESSAGE
        .Formula1 = TIME_LIST
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = g_UserOptions.DisplayEntryTips
        .ShowError = False
    End With

    With evalDateValidation
        .TypeOfValidation = xlValidateInputOnly
        .AlertStyle = xlValidAlertStop
        .InputTitle = DATE_INPUT_TITLE
        .InputMessage = dateInputMessage
        .ShowInput = g_UserOptions.DisplayEntryTips
        .ShowError = False
    End With

    With englishNameValidation
        .TypeOfValidation = xlValidateInputOnly
        .AlertStyle = xlValidAlertStop
        .InputTitle = STUDENT_ENG_NAME_TITLE
        .InputMessage = STUDENT_ENG_NAME_MESSAGE
        .ShowInput = g_UserOptions.DisplayEntryTips
        .ShowError = False
    End With

    With koreanNameValidation
        .TypeOfValidation = xlValidateInputOnly
        .AlertStyle = xlValidAlertStop
        .InputTitle = STUDENT_KOR_NAME_TITLE
        .InputMessage = STUDENT_KOR_NAME_MESSAGE
        .ShowInput = g_UserOptions.DisplayEntryTips
        .ShowError = False
    End With

    With studentGradeValidation
        .TypeOfValidation = xlValidateInputOnly
        .AlertStyle = xlValidAlertStop
        .InputTitle = STUDENT_GRADE_TITLE
        .InputMessage = STUDENT_GRADE_MESSAGE
        .ShowInput = g_UserOptions.DisplayEntryTips
        .ShowError = False
    End With

    With commentsValidation
        .TypeOfValidation = xlValidateInputOnly
        .AlertStyle = xlValidAlertStop
        .InputTitle = COMMENTS_TITLE
        .InputMessage = COMMENTS_MESSAGE
        .ShowInput = g_UserOptions.DisplayEntryTips
        .ShowError = False
    End With

    With teacherNotesValidation
        .TypeOfValidation = xlValidateInputOnly
        .AlertStyle = xlValidAlertStop
        .InputTitle = vbNullString
        .InputMessage = vbNullString
        .ShowInput = False
        .ShowError = False
    End With

    With ws
        ApplyValidationValues .Range(g_NATIVE_TEACHER), engTeacherValidation
        ApplyValidationValues .Range(g_KOREAN_TEACHER), korTeacherValidation
        ApplyValidationValues .Range(g_CLASS_LEVEL), classLevelValidation
        ApplyValidationValues .Range(g_CLASS_DAYS), classDaysValidation
        ApplyValidationValues .Range(g_CLASS_TIME), classTimeValidation
        ApplyValidationValues .Range(g_EVALUATION_DATE), evalDateValidation
        ApplyValidationValues .Range(g_ENGLISH_NAMES), englishNameValidation
        ApplyValidationValues .Range(g_KOREAN_NAMES), koreanNameValidation
        ApplyValidationValues .Range(g_STUDENT_GRADES), studentGradeValidation
        ApplyValidationValues .Range(g_COMMENTS), commentsValidation
        ApplyValidationValues .Range(g_TEACHER_NOTES), teacherNotesValidation
        
        ' Mind the colons (beginning and end of an extended range) and commas (next item in the list) in the range strings below
        ApplyFontSettingsToRange .Range(g_NATIVE_TEACHER & "," & g_CLASS_LEVEL & "," & g_CLASS_TIME & ":" & g_EVALUATION_DATE & "," & g_COMMENTS & ":" & g_TEACHER_NOTES), ENGLISH_FONT, 14, False
        ApplyFontSettingsToRange .Range(g_KOREAN_TEACHER), availablekoreanFont, 14, False
        ApplyFontSettingsToRange .Range(g_ENGLISH_NAMES), ENGLISH_FONT, 18, False
        ApplyFontSettingsToRange .Range(g_KOREAN_NAMES), availablekoreanFont, 20, False
        ApplyFontSettingsToRange .Range(g_STUDENT_GRADES), ENGLISH_FONT, 22, True
    End With
End Sub

Public Sub ToggleValidationTips()
    Dim ws As Worksheet
    Dim shtCodeName As String
    Dim fullTargetRange As String
    
    ' Mind the colons (beginning and end of an extended range) and commas (next item in the list) in the range strings below
    fullTargetRange = g_NATIVE_TEACHER & ":" & g_EVALUATION_DATE & "," & g_ENGLISH_NAMES & ":" & g_COMMENTS
    
    For Each ws In ThisWorkbook.Worksheets
        With ws
            shtCodeName = .CodeName
            Select Case shtCodeName
                Case "Instructions", "Options", "MacOS_Users", "Class_"
                    ' Hacky, but it looks clean and works.
                Case Else
                    ws.Range(fullTargetRange).Validation.ShowInput = g_UserOptions.DisplayEntryTips
            End Select
        End With
    Next ws
End Sub

Public Sub ApplyValidationValues(ByVal targetRange As Range, ByRef validatedValues As ValidationSettings)
    With targetRange.Validation
        On Error Resume Next
        .Delete
        On Error GoTo 0

        Select Case validatedValues.TypeOfValidation
            Case xlValidateList
                ' .Add Type:=validatedValues.TypeOfValidation, _
                     AlertStyle:=validatedValues.AlertStyle, _
                     Operator:=validatedValues.Operator, _
                     Formula1:=validatedValues.Formula1
                .Add Type:=validatedValues.TypeOfValidation, _
                     AlertStyle:=validatedValues.AlertStyle, _
                     Formula1:=validatedValues.Formula1
                .IgnoreBlank = validatedValues.IgnoreBlank
                .InCellDropdown = validatedValues.InCellDropdown
                .ShowInput = validatedValues.ShowInput
                .ShowError = validatedValues.ShowError
            Case Else
                .Add Type:=validatedValues.TypeOfValidation, _
                     AlertStyle:=validatedValues.AlertStyle
                .InputTitle = validatedValues.InputTitle
                .InputMessage = validatedValues.InputMessage
                .ShowInput = validatedValues.ShowInput
                .ShowError = validatedValues.ShowError
        End Select
    End With
End Sub

Private Sub ApplyFontSettingsToRange(ByVal targetRange As Range, ByVal fntName As String, ByVal fontSize As Long, ByVal fontBold As Boolean)
    With targetRange.Font
        .Name = fntName
        .Size = fontSize
        .Bold = fontBold
        .Italic = False
        .Underline = False
    End With
End Sub

Public Sub ToggleEmbeddedSignature(ByVal clickedButtonName As String)
    Dim optionShps As Shapes: Set optionShps = Options.Shapes
    
    With optionShps
        On Error Resume Next
        .Item("Button_SignatureMissing").Visible = Not optionShps("Button_SignatureMissing").Visible
        .Item("Button_SignatureEmbedded").Visible = Not optionShps("Button_SignatureEmbedded").Visible
        On Error GoTo 0
        
        Select Case clickedButtonName
            Case "Button_SignatureMissing"
                .Item("mySignature-Disabled").Name = "mySignature-Enabled"
            Case "Button_SignatureEmbedded"
                .Item("mySignature-Enabled").Name = "mySignature-Disabled"
        End Select
    End With
End Sub

#If Mac Then
Public Sub SetVisibilityOfMacSettingsShapes(ByVal isAppleScriptInstalled As Boolean, ByVal isDialogToolkitInstalled As Boolean)
    Dim enhancedDialogsStatus      As String
    Dim enhancedDialogsAreDisabled As Boolean

    With MacOS_Users.Shapes
        .Item("Button_SpeakingEvalsScpt_Missing").Visible = Not isAppleScriptInstalled
        .Item("Button_DialogToolkit_Missing").Visible = Not isDialogToolkitInstalled
        .Item("Button_SpeakingEvalsScpt_Installed").Visible = isAppleScriptInstalled
        .Item("Button_DialogToolkit_Installed").Visible = isDialogToolkitInstalled
        
        enhancedDialogsStatus = MacOS_Users.Cells(1, 1).Value
        enhancedDialogsAreDisabled = Not isDialogToolkitInstalled _
                                     Or enhancedDialogsStatus = "Enhanced Dialogs: Disabled" _
                                     Or enhancedDialogsStatus = vbNullString
        
        .Item("Button_EnhancedDialogs_Disable").Visible = enhancedDialogsAreDisabled
        .Item("Button_EnhancedDialogs_Enable").Visible = Not enhancedDialogsAreDisabled
    End With
    
    If enhancedDialogsAreDisabled And enhancedDialogsStatus <> "Enhanced Dialogs: Disabled" Then
        ToggleSheetProtection MacOS_Users, False
        WriteNewRangeValue MacOS_Users.Cells(1, 1), "Enhanced Dialogs: Disabled"
        ToggleSheetProtection MacOS_Users, True
    End If
End Sub

Public Sub ToggleMacSettingsButtons(ByRef ws As Worksheet, ByVal clickedButtonName As String)
    Const SCRIPT_ENABLED  As String = "Enhanced Dialogs: Enabled"
    Const SCRIPT_DISABLED As String = "Enhanced Dialogs: Disabled"
    
    Dim shps As Shapes
    Dim installedStatus As Boolean

    Set shps = ws.Shapes

    If shps("Button_DialogToolkit_Missing").Visible Then
        installedStatus = AreAppleScriptsInstalled(, , True)
        
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
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.Workbook.UpdatingPersistantSetting")
    End If
    
    ToggleSheetProtection ws, False
    With ws
        WriteNewRangeValue .Cells(1, 1), IIf(.Shapes("Button_EnhancedDialogs_Enable").Visible, SCRIPT_ENABLED, SCRIPT_DISABLED)

        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.Workbook.UpdatingPersistantSettingValue", .Cells(1, 1).Value)
        End If
    End With
    ToggleSheetProtection ws, True
End Sub
#End If