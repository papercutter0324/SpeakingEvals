Option Explicit

#Const PRINT_DEBUG_MESSAGES = True
#If Mac Then
    Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
    Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
#End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Worksheet Layout and Formatting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub VerifySheetNames()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Sheets
        With ws
            Select Case .CodeName
                Case "Instructions"
                    If .Name <> "Instructions" Then .Name = "Instructions"
                Case "MacOS_Users"
                    If .Name <> "MacOS Users" Then .Name = "MacOS Users"
                Case "Options"
                    If .Name <> "Options" Then .Name = "Options"
            End Select
        End With
    Next ws
End Sub

Public Sub AutoPopulateEvaluationDateValues(ByVal ws As Worksheet)
    Dim dateCell As Range
    Dim dateAsDate As Date
    Dim dateToCheck As Date
    Dim messageText As String
    Dim msgResult As Variant
    
    On Error Resume Next
    If ws.Range("A6").Value = "Evaluation Date:" Then
        Set dateCell = ws.Range("C6")

        If Len(Trim$(dateCell.Value)) = 0 Then
            dateCell.Value = Format$(Date, "DD MMM. YYYY")
        ElseIf IsDate(Trim$(dateCell.Value)) Then
            ' This isn't working
            
            ' dateAsDate = CDate(Trim$(dateCell.Value))
            ' dateToCheck = DateAdd("m", -2, Date)
            
            ' If dateAsDate < dateToCheck Then
            '     dateCell.Value = Format$(Date, "MMM. YYYY")
            ' End If
        Else
            messageText = "An invalid date has been found on worksheet " & ws.Name & "." & vbNewLine & _
                          "Please enter a valid date."
            msgResult = DisplayMessage(messageText, vbInformation, "Invalid Date!")
            dateCell.Value = vbNullString
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub SetLayoutInstructions()
    Dim shp As Shapes
    Dim shapeProps As Variant
    Dim buttonNames As Variant
    Dim i As Long
    
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
    
    buttonNames = Array("Button_Speadsheet", "Button_Font", "Button_ReportTemplate", "Button_SignatureTemplate", "Button_SourceCode")
    
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
    
    Set shp = Instructions.Shapes
    
    ' Loop through the shape properties array and apply the settings
    On Error Resume Next
    For i = LBound(shapeProps) To UBound(shapeProps)
        With shp.Item(shapeProps(i)(0))
            If Err.Number = 0 Then
                .Top = shapeProps(i)(1)
                .Left = shapeProps(i)(2)
                .Height = shapeProps(i)(3)
                .Width = shapeProps(i)(4)
            Else
                #If PRINT_DEBUG_MESSAGES Then
                    Debug.Print INDENT_LEVEL_2 & "Warning: Shape '" & shapeProps(i)(0) & "' not found."
                #End If
                Err.Clear ' Clear the error
            End If
        End With
    Next i
    
    ' Position buttons
    For i = LBound(buttonNames) To UBound(buttonNames)
        With shp.Item(buttonNames(i))
            If Err.Number = 0 Then
                .Top = BUTTON_TOP
                .Left = CODE_MSG_LEFT + 20 + i * (BUTTON_WIDTH + CELL_SPACING)
                .Height = BUTTON_HEIGHT
                .Width = BUTTON_WIDTH
            Else
                #If PRINT_DEBUG_MESSAGES Then
                    Debug.Print INDENT_LEVEL_2 & "Warning: Shape '" & buttonNames(i) & "' not found."
                #End If
                Err.Clear ' Clear the error
            End If
        End With
    Next i
    On Error GoTo 0
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print INDENT_LEVEL_2 & "Result: " & IIf(Err.Number = 0, "Complete", "Errors found.")
    #End If
End Sub

Public Sub SetLayoutMacOSUsers()
    Dim shp As Shapes
    Dim shapeProps As Variant
    Dim buttonProps As Variant
    Dim i As Long
    
    Const MACOS_TB_TOP As Double = 15
    Const MACOS_TB_LEFT As Double = 15
    Const MACOS_TB_HEIGHT As Double = 58
    Const MACOS_TB_WIDTH As Double = 1285
    
    Const MACOS_MSG_TOP As Double = MACOS_TB_TOP + MACOS_TB_HEIGHT
    Const MACOS_MSG_LEFT As Double = MACOS_TB_LEFT
    Const MACOS_MSG_HEIGHT As Double = 800
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
    
    Set shp = MacOS_Users.Shapes
    
    ' Loop through the shape properties array and apply the settings
    On Error Resume Next
    For i = LBound(shapeProps) To UBound(shapeProps)
        With shp.Item(shapeProps(i)(0))
            If Err.Number = 0 Then
                .Top = shapeProps(i)(1)
                .Left = shapeProps(i)(2)
                .Height = shapeProps(i)(3)
                .Width = shapeProps(i)(4)
            Else
                #If PRINT_DEBUG_MESSAGES Then
                    Debug.Print INDENT_LEVEL_2 & "Warning: Shape '" & shapeProps(i)(0) & "' not found."
                #End If
                Err.Clear ' Clear the error
            End If
        End With
    Next i

    ' Loop through button properties and set positions
    For i = LBound(buttonProps) To UBound(buttonProps)
        With shp.Item(buttonProps(i)(0))
            If Err.Number = 0 Then
                .Top = BUTTON_TOP
                .Left = buttonProps(i)(1)
                .Height = BUTTON_HEIGHT
                .Width = BUTTON_WIDTH
            Else
                #If PRINT_DEBUG_MESSAGES Then
                    Debug.Print INDENT_LEVEL_2 & "Warning: Shape '" & buttonProps(i)(0) & "' not found."
                #End If
                Err.Clear ' Clear the error
            End If
        End With
    Next i
    On Error GoTo 0
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print INDENT_LEVEL_2 & "Result: " & IIf(Err.Number = 0, "Complete", "Errors found.")
    #End If
End Sub

Public Sub SetLayoutOptions()
    Dim shp As Shapes
    Dim shapeProps As Variant
    Dim aspectRatio As Double
    Dim mySignatureCalculatedHeight As Double
    Dim mySignatureCalculatedWidth As Double
    Dim mySignatureCalculatedTop As Double
    Dim mySignatureCalculatedLeft As Double
    Dim i As Long
    
    ' Signature Information Title Bar
    Const TB_HEIGHT As Double = 58
    Const TB_WIDTH As Double = 1030
    Const TB_TOP As Double = 15
    Const TB_LEFT As Double = 15
    
    ' Signature Image Title  Bar
    Const SIG_TB_HEIGHT As Double = TB_HEIGHT
    Const SIG_TB_WIDTH As Double = 300
    Const SIG_TB_TOP As Double = TB_TOP
    Const SIG_TB_LEFT As Double = TB_LEFT + TB_WIDTH
    
    ' Signature Information Msg Box
    Const MSG_HEIGHT As Double = 320
    Const MSG_WIDTH As Double = TB_WIDTH + SIG_TB_WIDTH
    Const MSG_TOP As Double = TB_TOP + TB_HEIGHT
    Const MSG_LEFT As Double = TB_LEFT
    
    ' Signature Image Container
    Const SIG_CONTAINER_HEIGHT As Double = 86
    Const SIG_CONTAINER_WIDTH As Double = SIG_TB_WIDTH
    Const SIG_CONTAINER_TOP As Double = SIG_TB_TOP + SIG_TB_HEIGHT
    Const SIG_CONTAINER_LEFT As Double = SIG_TB_LEFT
    
    ' Signature Image Dimensions
    Const MAX_HEIGHT As Double = 68
    Const MAX_WIDTH As Double = 286
    Const DEFAULT_ASPECT_RATIO As Double = MAX_WIDTH / MAX_HEIGHT
    
    ' Signature Image
    Const SIG_PLACEHOLDER_HEIGHT As Double = MAX_HEIGHT
    Const SIG_PLACEHOLDER_WIDTH As Double = MAX_WIDTH
    Const SIG_PLACEHOLDER_TOP As Double = SIG_CONTAINER_TOP + (SIG_CONTAINER_HEIGHT - SIG_PLACEHOLDER_HEIGHT) / 2
    Const SIG_PLACEHOLDER_LEFT As Double = SIG_CONTAINER_LEFT + (SIG_CONTAINER_WIDTH - SIG_PLACEHOLDER_WIDTH) / 2
    
    ' Signature Toggle Button Container
    Const BTN_CONTAINER_HEIGHT As Double = SIG_CONTAINER_HEIGHT
    Const BTN_CONTAINER_WIDTH As Double = 205
    Const BTN_CONTAINER_TOP As Double = SIG_CONTAINER_TOP + SIG_CONTAINER_HEIGHT
    Const BTN_CONTAINER_LEFT As Double = SIG_CONTAINER_LEFT + SIG_CONTAINER_WIDTH - (SIG_CONTAINER_WIDTH / 2) - (BTN_CONTAINER_WIDTH / 2)
    
    ' Signature Toggle Button
    Const BTN_SIGNATURE_HEIGHT As Double = 65
    Const BTN_SIGNATURE_WIDTH As Double = 175
    Const BTN_SIGNATURE_TOP As Double = SIG_CONTAINER_TOP + SIG_CONTAINER_HEIGHT + 10
    Const BTN_SIGNATURE_LEFT As Double = SIG_CONTAINER_LEFT + (SIG_CONTAINER_WIDTH - BTN_SIGNATURE_WIDTH) / 2
    
    ' Winner Certificates Title Bar
    Const TB_CERTS_HEIGHT As Double = TB_HEIGHT
    Const TB_CERTS_WIDTH As Double = 620
    Const TB_CERTS_TOP As Double = MSG_TOP + MSG_HEIGHT + 15
    Const TB_CERTS_LEFT As Double = TB_LEFT
    
    ' Winner Certificates Options and Preview Title Bar
    Const TB_CERT_OPTIONS_HEIGHT As Double = TB_HEIGHT
    Const TB_CERT_OPTIONS_WIDTH As Double = TB_WIDTH + SIG_TB_WIDTH - TB_CERTS_WIDTH
    Const TB_CERT_OPTIONS_TOP As Double = TB_CERTS_TOP
    Const TB_CERT_OPTIONS_LEFT As Double = TB_TOP + TB_CERTS_WIDTH
    
    ' Winner Certificates Information Box
    Const MSG_CERTS_HEIGHT As Double = 340
    Const MSG_CERTS_WIDTH As Double = TB_CERTS_WIDTH
    Const MSG_CERTS_TOP As Double = TB_CERTS_TOP + TB_CERTS_HEIGHT
    Const MSG_CERTS_LEFT As Double = TB_CERTS_LEFT
    
    ' Winner Certificates Preview Images
    Const CERT_PREVIEW_HEIGHT As Double = 300
    Const CERT_PREVIEW_WIDTH As Double = 433.32
    Const CERT_PREVIEW_TOP As Double = MSG_CERTS_TOP + (MSG_CERTS_HEIGHT / 2) - (CERT_PREVIEW_HEIGHT / 2)
    Const CERT_PREVIEW_LEFT As Double = TB_CERT_OPTIONS_LEFT + 20
    

    ' Array of the standard shapes and their positions/dimensions
    shapeProps = Array( _
        Array("Title Bar", TB_TOP, TB_LEFT, TB_HEIGHT, TB_WIDTH), _
        Array("Message", MSG_TOP, MSG_LEFT, MSG_HEIGHT, MSG_WIDTH), _
        Array("Signature Title Bar", SIG_TB_TOP, SIG_TB_LEFT, SIG_TB_HEIGHT, SIG_TB_WIDTH), _
        Array("Signature Container", SIG_CONTAINER_TOP, SIG_CONTAINER_LEFT, SIG_CONTAINER_HEIGHT, SIG_CONTAINER_WIDTH), _
        Array("Button_Container", BTN_CONTAINER_TOP, BTN_CONTAINER_LEFT, BTN_CONTAINER_HEIGHT, BTN_CONTAINER_WIDTH), _
        Array("Button_SignatureEmbedded", BTN_SIGNATURE_TOP, BTN_SIGNATURE_LEFT, BTN_SIGNATURE_HEIGHT, BTN_SIGNATURE_WIDTH), _
        Array("Button_SignatureMissing", BTN_SIGNATURE_TOP, BTN_SIGNATURE_LEFT, BTN_SIGNATURE_HEIGHT, BTN_SIGNATURE_WIDTH), _
        Array("Certificate_TitleBar", TB_CERTS_TOP, TB_CERTS_LEFT, TB_CERTS_HEIGHT, TB_CERTS_WIDTH), _
        Array("Certificate_Message", MSG_CERTS_TOP, MSG_CERTS_LEFT, MSG_CERTS_HEIGHT, MSG_CERTS_WIDTH), _
        Array("Certificate_Options_TitleBar", TB_CERT_OPTIONS_TOP, TB_CERT_OPTIONS_LEFT, TB_CERT_OPTIONS_HEIGHT, TB_CERT_OPTIONS_WIDTH), _
        Array("Preview_Design_Landscape_Default", CERT_PREVIEW_TOP, CERT_PREVIEW_LEFT, CERT_PREVIEW_HEIGHT, CERT_PREVIEW_WIDTH), _
        Array("Preview_Border_Landscape_Style 1", CERT_PREVIEW_TOP, CERT_PREVIEW_LEFT, CERT_PREVIEW_HEIGHT, CERT_PREVIEW_WIDTH), _
        Array("Preview_Border_Landscape_Style 2", CERT_PREVIEW_TOP, CERT_PREVIEW_LEFT, CERT_PREVIEW_HEIGHT, CERT_PREVIEW_WIDTH) _
    )

    Set shp = Options.Shapes

    ' Step 1: Verify all shapes exist
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print INDENT_LEVEL_2 & "Verifying all standard shapes are present."
    #End If
    For i = LBound(shapeProps) To UBound(shapeProps)
        If Not DoesShapeExist(Options, shapeProps(i)(0)) Then
            #If PRINT_DEBUG_MESSAGES Then
                Debug.Print INDENT_LEVEL_3 & "Warning: Shape '" & shapeProps(i)(0) & "' not found."
            #End If
            ' Exit Sub ' Support for this will be added later
        End If
        
        ' Step 2: Set dimensions and positions of standard shapes and textboxes
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print INDENT_LEVEL_2 & "Setting layout for " & shapeProps(i)(0) & "."
        #End If
        SetButtonDimensionsAndPosition shp.Item(shapeProps(i)(0)), shapeProps(i)(3), shapeProps(i)(4), _
                                       shapeProps(i)(1), shapeProps(i)(2)
    Next i

    ' Step 3a: Set dimensions and positions of 'mySignature_Placeholder', if present
    If DoesShapeExist(Options, "mySignature_Placeholder") Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print INDENT_LEVEL_2 & "'mySignature_Placeholder' found." & vbNewLine & _
                        INDENT_LEVEL_2 & "    Verifying correct dimensions."
        #End If
        SetButtonDimensionsAndPosition shp.Item("mySignature_Placeholder"), SIG_PLACEHOLDER_HEIGHT, _
                                       SIG_PLACEHOLDER_WIDTH, SIG_PLACEHOLDER_TOP, SIG_PLACEHOLDER_LEFT
    End If

    ' Step 3b: Set dimensions and positions of 'mySignature', if present
    If DoesShapeExist(Options, "mySignature") Then
        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print INDENT_LEVEL_2 & "'mySignature' found." & vbNewLine & _
                        INDENT_LEVEL_2 & "    Verifying correct dimensions."
        #End If
        
        mySignatureCalculatedHeight = MAX_HEIGHT
        mySignatureCalculatedWidth = MAX_WIDTH
        
        With shp.Item("mySignature")
            mySignatureCalculatedTop = SIG_CONTAINER_TOP + (SIG_CONTAINER_HEIGHT - .Height) / 2
            mySignatureCalculatedLeft = SIG_CONTAINER_LEFT + (SIG_CONTAINER_WIDTH - .Width) / 2
            aspectRatio = .Width / .Height
        End With
        
        If DEFAULT_ASPECT_RATIO > aspectRatio Then
            mySignatureCalculatedWidth = MAX_HEIGHT * aspectRatio
        Else
            mySignatureCalculatedHeight = MAX_WIDTH / aspectRatio
        End If
        
        SetButtonDimensionsAndPosition shp.Item("mySignature"), mySignatureCalculatedHeight, mySignatureCalculatedWidth, _
                                       mySignatureCalculatedTop, mySignatureCalculatedLeft
    End If
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print INDENT_LEVEL_2 & "Result: " & IIf(Err.Number = 0, "Complete", "Errors found.")
    #End If
End Sub

Public Sub OptionsShapeVisibility(ByVal ws As Worksheet)
    Dim shp As Shapes
    Dim defaultOption As String
    Dim signaturePresent As Boolean
    Dim certificateOptions As Variant
    Dim optionRangeRow As Long
    Dim i As Long
    
    Set shp = ws.Shapes
    
    ' Step 1: Verify correct signature button is displayed
    On Error Resume Next
    signaturePresent = Not shp.[_Default]("mySignature") Is Nothing
    On Error GoTo 0
    
    shp.[_Default]("Button_SignatureEmbedded").Visible = signaturePresent
    shp.[_Default]("Button_SignatureMissing").Visible = Not signaturePresent
    
    ' Step 2: Check certificate options, setting defaults as needed, and generate preview
    certificateOptions = ws.Range("J10:K14")
    For i = LBound(certificateOptions, 1) To UBound(certificateOptions, 1)
        If certificateOptions(i, 2) = vbNullString Then
            optionRangeRow = i + 9
            defaultOption = GetDefaultCertificateOptions(ws.Range("k" & optionRangeRow), certificateOptions(i, 1))
            ws.Range("k" & optionRangeRow).Value = defaultOption
        End If
    Next i
End Sub

Public Function GetDefaultCertificateOptions(ByVal certOption As Range, ByVal optionType As String) As String
    Select Case optionType
        Case "Layout:"
            GetDefaultCertificateOptions = "Landscape"
        Case "Design:"
            GetDefaultCertificateOptions = "Default"
        Case "Border:"
            GetDefaultCertificateOptions = "Disabled"
        Case "Border Color:"
            GetDefaultCertificateOptions = "Default"
        Case "Color Code:"
            GetDefaultCertificateOptions = GetCertificateBorderColorCode(certOption.Offset(-2, 0).Value, certOption.Offset(-1, 0).Value)
    End Select
End Function

Public Function GetCertificateBorderColorCode(ByVal borderStyle As String, ByVal borderColorOption As String) As String
    Select Case borderColorOption
        Case "Default"
            GetCertificateBorderColorCode = GetDefaultBorderColor(borderStyle)
        Case "Gold"
            GetCertificateBorderColorCode = "#EFBF04"
        Case "Metalic Gold"
            GetCertificateBorderColorCode = "#D4AF37"
        Case "Silver"
            GetCertificateBorderColorCode = "#C0C0C0"
        Case "Dark Teal"
            GetCertificateBorderColorCode = "#2B694A"
        Case "Custom"
            GetCertificateBorderColorCode = GetDefaultBorderColor(borderStyle)
    End Select
End Function

Public Function GetCertificateBorderColorLabel(ByVal borderColorCode As String) As String
    Select Case borderColorCode
        Case "#EFBF04"
            GetCertificateBorderColorLabel = "Gold"
        Case "#D4AF37"
            GetCertificateBorderColorLabel = "Metalic Gold"
        Case "#C0C0C0"
            GetCertificateBorderColorLabel = "Silver"
        Case "#2B694A"
            GetCertificateBorderColorLabel = "Dark Teal"
        Case Else
            GetCertificateBorderColorLabel = "Custom"
    End Select
End Function

Private Function GetDefaultBorderColor(ByVal borderStyle As String) As String
    Select Case borderStyle
        Case "Disabled"
            GetDefaultBorderColor = "#000000"
        Case "Style 1"
            GetDefaultBorderColor = "#EFBF04"
        Case "Style 2"
            GetDefaultBorderColor = "#2B694A"
        Case Else
            GetDefaultBorderColor = "#EFBF04"
    End Select
End Function

Private Function DoesShapeExist(ByVal ws As Worksheet, ByVal shapeName As String) As Boolean
    On Error Resume Next
    DoesShapeExist = Not ws.Shapes.[_Default](shapeName) Is Nothing
    On Error GoTo 0
End Function

Private Sub SetButtonDimensionsAndPosition(ByVal buttonShape As Shape, ByVal buttonHeight As Double, ByVal buttonWidth As Double, ByVal buttonTop As Double, ByVal buttonLeft As Double)
    With buttonShape
        .LockAspectRatio = msoFalse
        .Height = buttonHeight
        .Width = buttonWidth
        .LockAspectRatio = msoTrue
        .Top = buttonTop
        .Left = buttonLeft
    End With
End Sub

Public Sub RepairLayouts(ByVal ws As Worksheet)
    ToggleSheetProtection ws, False
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Repairing Layout" & vbNewLine & _
                    INDENT_LEVEL_1 & "Sheet: " & ws.Name
    #End If
    
    SetLayoutClassRecords ws
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print INDENT_LEVEL_1 & "Repair Complete"
    #End If
    
    ToggleSheetProtection ws, True
End Sub

Public Sub SetLayoutClassRecords(ByVal ws As Worksheet)
    Dim borderRangeA As Range
    Dim borderRangeB As Range
    Dim existingName As Name
    Dim shadingRanges As Variant
    Dim i As Long

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
        Array("buttonShadingRange", "J1:J6", RGB(255, 255, 255)), _
        Array("commentHeaderShadingRange", "J7", RGB(191, 191, 191)), _
        Array("commentValuesShadingRange", "J8:J32", RGB(242, 242, 242)), _
        Array("notesHeaderShadingRange", "K1,K2:M6,K7", RGB(196, 189, 151)), _
        Array("notesValuesShadingRange", "K8:M32", RGB(221, 217, 196)) _
    )

    ' Step 1: Apply Bulk Formatting & Named Ranges
    With ws
        ' Set Column widths and Row heights
        With .Columns
            .Item("A").ColumnWidth = 7
            .Item("B:C").ColumnWidth = 18
            .Item("D:I").ColumnWidth = 22
            .Item("J").ColumnWidth = 103
            .Item("K").ColumnWidth = 20
            .Item("L").ColumnWidth = 50
            .Item("M").ColumnWidth = 10
        End With
        With .Rows
            .Item("1:6").RowHeight = 30
            .Item("7").RowHeight = 25
            .Item("8:32").RowHeight = 50
        End With

        ' Set Colors and Named Ranges
        With ThisWorkbook.names
            For i = LBound(shadingRanges) To UBound(shadingRanges)
                On Error Resume Next
                Set existingName = .Item(shadingRanges(i)(0))
                If Not existingName Is Nothing Then existingName.Delete
                On Error GoTo 0

                .Add Name:=shadingRanges(i)(0), RefersTo:=ws.Range(shadingRanges(i)(1)) ' Add the new named range
            Next i
        End With
        
        For i = LBound(shadingRanges) To UBound(shadingRanges)
            .Range(shadingRanges(i)(1)).Interior.Color = shadingRanges(i)(2) ' Set background color
        Next i

        ' Set Bulk Borders
        Set borderRangeA = .Range("A1:C6,A8:A32,B8:C32,D8:I32,J8:J32,K8:M32") ' Main data areas
        Set borderRangeB = .Range("D1:I6,J1:J6,K1:M6") ' Header/Button areas

        With borderRangeA.Borders ' Apply to all borders first
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Item(xlInsideHorizontal).LineStyle = xlDash
            .Item(xlInsideHorizontal).Weight = xlThin
            .Item(xlInsideVertical).LineStyle = xlLineStyleNone
        End With
        With borderRangeB.Borders ' No inside borders for these areas
            .Item(xlInsideHorizontal).LineStyle = xlLineStyleNone
            .Item(xlInsideVertical).LineStyle = xlLineStyleNone
        End With

        With .Range("C1:C6,B8:M32,L2:L4")
            .Locked = False
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            .WrapText = True
            .NumberFormat = "@" ' Treat as text
        End With
        .Range("A1:B6,A7:M7,D1:J6,K1:M1,K2:K6,L5:L6,M2:M6,A8:A32").Locked = True
    End With

    ' Step 2: Position Buttons
    SetLayoutClassRecordsButtons ws

    ' Step 3: Apply Validation and Specific Formatting
    VerifyValidationAndFontsForClassRecords ws
    
    ' Step 4: Generate winner list
    GenerateCompleteHiddenNameValidationList ws
    PopulateWinnersListValidationValues ws
    
    ' Step 5: Set shading
    SetDefaultShading ws

    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print INDENT_LEVEL_2 & "Result: " & IIf(Err.Number = 0, "Complete", "Errors found.")
    #End If
End Sub

Public Sub SetLayoutClassRecordsButtons(ByVal ws As Worksheet)
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
            With .Item(buttonProps(i)(0))
                .Height = BUTTON_HEIGHT
                .Width = BUTTON_WIDTH
                .Top = cellTop + (cellVerticalSpacing * buttonProps(i)(1)) + ((buttonProps(i)(1) - 1) * BUTTON_HEIGHT)
                .Left = cellLeft + (cellHorizontalSpacing * buttonProps(i)(2)) + ((buttonProps(i)(2) - 1) * BUTTON_WIDTH)
            End With
        Next i
    End With
End Sub

Private Sub VerifyValidationAndFontsForClassRecords(ByVal ws As Worksheet)
    Dim dateInputMessage As String
    Dim currentFont As String
    Dim koreanFont As String
    Dim validationList As String
    Dim validationValues As Variant
    Dim i As Long
    Dim cellRng As Range
    
    Const ENGLISH_FONT As String = "Calibri"
    Const KOREAN_FONT_DEFAULT As String = "Malgun Gothic"
    Const KOREAN_FONT_CUSTOM As String = "Kakao Big Sans"
    Const KOREAN_FONT_CUSTOM_FILENAME As String = "KakaoBigSans-Regular.ttf"
    
    Const LEVEL_LIST As String = "Theseus,Perseus,Odysseus,Hercules,Artemis,Hermes,Apollo,Zeus,E5 Athena,Helios,Poseidon,Gaia,Hera,E6 Song's"
    Const DAYS_LIST As String = "MonWed,MonFri,WedFri,MWF,TTh,MWF (Class 1),MWF (Class 2),TTh (Class 1),TTh (Class 2)"
    Const TIME_LIST As String = "9pm,830pm,8pm,7pm,6pm,530pm,5pm,4pm,3pm,2pm,1pm,12pm,11am,10am,9am"
    
    Const ENGLISH_NAME_RANGE As String = "B8:B32"
    Const KOREAN_NAME_RANGE As String = "C8:C32"
    Const GRADES_RANGE As String = "D8:I32"
    Const COMMENTS_RANGE As String = "J8:J32"
    Const NOTES_RANGE As String = "K8:K32"
    
    #If Mac Then
        ' No extra variables needed
    #Else
        Dim fso As Object
        Dim userFontPath As String
        Dim sysFontPath As String
    #End If
    
    ' Step 1: Determine Korean font to be used
    #If Mac Then
        ' Figure out how to check...
        koreanFont = KOREAN_FONT_DEFAULT
    #Else
        Set fso = CreateObject("Scripting.FileSystemObject")
        userFontPath = fso.BuildPath(Environ$("LOCALAPPDATA") & "\Microsoft\Windows\Fonts", KOREAN_FONT_CUSTOM_FILENAME)
        sysFontPath = fso.BuildPath(Environ$("WINDIR") & "\Fonts", KOREAN_FONT_CUSTOM_FILENAME)
        
        If fso.fileExists(userFontPath) Or fso.fileExists(sysFontPath) Then
            koreanFont = KOREAN_FONT_CUSTOM
        Else
            koreanFont = KOREAN_FONT_DEFAULT
        End If
    #End If

    ' Step 2: Detect date format
    Select Case Application.International(xlDateOrder)
       Case 0: dateInputMessage = "MM/DD/YYYY" & vbNewLine & "or MM/YYYY."
       Case 1: dateInputMessage = "DD/MM/YYYY" & vbNewLine & "or MM/YYYY."
       Case 2: dateInputMessage = "YYYY/MM/DD" & vbNewLine & "or MM/YYYY."
    End Select

    ' Step 3: Prepare and apply validation values specific to C1:C6
    validationValues = Array( _
        Array("Native Teacher's Name", "Please enter just your" & vbNewLine & "name, no suffix or title" & vbNewLine & "like ""tr.""", vbNullString), _
        Array("Korean Teacher's Name", "Please write their Korean name. The parents are unlikely to know their English name.", vbNullString), _
        Array("Class Level", "Click on the down arrow and choose the class's level from the list.", LEVEL_LIST), _
        Array("Class Days", "Select the days when you see this class." & vbNewLine & vbNewLine & "For Athena and Song's classes, use Class-1 and Class-2 to help organize split classes.", DAYS_LIST), _
        Array("Class Time", "Select what time you have Class 1 each week. Scroll to see more options." & vbNewLine & vbNewLine & "This is to help you keep track of which class this is; it won't appear on the final reports.", TIME_LIST), _
        Array("Date Format", dateInputMessage, vbNullString) _
    )

    ApplyValidationAndFontsForClassRecords ws, ws.Range("C1,C6"), ENGLISH_FONT, 14, False, xlValidateInputOnly, xlValidAlertStop, vbNullString, vbNullString, True, False
    ApplyValidationAndFontsForClassRecords ws, ws.Range("C2"), koreanFont, 14, False, xlValidateInputOnly, xlValidAlertStop, vbNullString, vbNullString, True, False
    ApplyValidationAndFontsForClassRecords ws, ws.Range("C3:C5"), ENGLISH_FONT, 14, False, xlValidateList, xlValidAlertStop, vbNullString, vbNullString, True, False
    
    ' Step 3a: Handle special exceptions for C1:C6
    For i = 1 To 6
        Set cellRng = ws.Cells.Item(i, 3)
        With cellRng.Validation
            If i >= 3 And i <= 5 Then
                On Error Resume Next
                .Delete
                On Error GoTo 0
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=validationValues(i - 1)(2)
                .ShowInput = True
                .ShowError = False
            End If
            .InputTitle = validationValues(i - 1)(0)
            .InputMessage = validationValues(i - 1)(1)
        End With
    Next i

    ' Step 4: Apply validation values for the remaining sections
    ApplyValidationAndFontsForClassRecords ws, ws.Range(ENGLISH_NAME_RANGE), ENGLISH_FONT, 18, False, xlValidateInputOnly, xlValidAlertStop, "Character Limit", "20 characters or fewer recommended", True, False
    ApplyValidationAndFontsForClassRecords ws, ws.Range(KOREAN_NAME_RANGE), koreanFont, 20, False, xlValidateInputOnly, xlValidAlertStop, "Language Reminder", "Please write their names in Korean.", True, False
    ApplyValidationAndFontsForClassRecords ws, ws.Range(GRADES_RANGE), ENGLISH_FONT, 22, True, xlValidateInputOnly, xlValidAlertStop, "Enter a Grade", "Valid Letter Grades" & vbNewLine & "  A+ / A / B+ / B / C" & vbNewLine & vbNewLine & "Valid Numeric Scores   " & vbNewLine & "  1 ~ 5", True, False
    ApplyValidationAndFontsForClassRecords ws, ws.Range(COMMENTS_RANGE), ENGLISH_FONT, 14, False, xlValidateInputOnly, xlValidAlertStop, "Character Limit", "960 characters", True, False
    ApplyValidationAndFontsForClassRecords ws, ws.Range(NOTES_RANGE), ENGLISH_FONT, 14, False, xlValidateInputOnly, xlValidAlertStop, vbNullString, vbNullString, False, False
End Sub

Private Sub ApplyValidationAndFontsForClassRecords(ByVal ws As Worksheet, ByVal currentRange As Range, ByVal fontName As String, ByVal fontSize As Long, ByVal fontBold As Boolean, _
                                                   ByVal valType As Variant, ByVal valAlertStyle As Variant, ByVal valInputTitle As String, ByVal valInputMsg As String, ByVal valShowInput As Boolean, ByVal valShowError As Boolean)
    With currentRange
        ' Step 1: Set font settings
        With .Font
            .Name = fontName
            .Size = fontSize
            .Bold = fontBold
            .Italic = False
            .Underline = False
        End With
        
        ' Exit early if a validation list is used for the current range
        If valType = xlValidateList Then Exit Sub
        
        ' Step 2: Set validation settings
        On Error Resume Next
        .Validation.Delete
        On Error GoTo 0
        
        With .Validation
            .Add Type:=valType, AlertStyle:=valAlertStyle
            .InputTitle = valInputTitle
            .InputMessage = valInputMsg
            .ShowInput = valShowInput
            .ShowError = valShowError
        End With
    End With
End Sub

Public Sub ToggleEmbeddedSignature(ByVal clickedButtonName As String)
    Dim shp As Shapes
    
    Set shp = Options.Shapes
    
    On Error Resume Next
    shp.[_Default]("Button_SignatureMissing").Visible = Not shp.[_Default]("Button_SignatureMissing").Visible
    shp.[_Default]("Button_SignatureEmbedded").Visible = Not shp.[_Default]("Button_SignatureEmbedded").Visible
    
    Select Case clickedButtonName
        Case "Button_SignatureMissing"
            shp.[_Default]("mySignature_Placeholder").Name = "mySignature"
        Case "Button_SignatureEmbedded"
            shp.[_Default]("mySignature").Name = "mySignature_Placeholder"
    End Select
    On Error GoTo 0
    
    Set shp = Nothing
End Sub

#If Mac Then
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
        ToggleSheetProtection ws, False
        ws.Cells(1, 1).Value = "Enhanced Dialogs: Disabled"
        ToggleSheetProtection ws, True
    End If
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
        Debug.Print INDENT_LEVEL_1 & "Updating persistant status value."
    #End If
    
    ToggleSheetProtection ws, False
    With ws
        .Cells(1, 1).Value = IIf(.Shapes("Button_EnhancedDialogs_Enable").Visible, SCRIPT_ENABLED, SCRIPT_DISABLED)

        #If PRINT_DEBUG_MESSAGES Then
            Debug.Print INDENT_LEVEL_1 & "Value: """ & .Cells(1, 1).Value & """"
        #End If
    End With
    ToggleSheetProtection ws, True
End Sub
#End If
