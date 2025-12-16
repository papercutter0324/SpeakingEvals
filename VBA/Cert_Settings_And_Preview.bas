Option Explicit

#Const Windows = (Mac = 0)

Public Type CertificateDesign
    Type                As String
    Layout              As String
    Design              As String
    borderType          As String
    BorderColor         As String
    borderColorCode     As String
End Type

Public Sub UpdateCertificateDesign(ByVal updatedCellsRange As Range)
    UpdateSubSettings updatedCellsRange
    UpdateCertificatePreview
End Sub

Private Sub UpdateSubSettings(ByRef updatedRange As Range)
    Dim targetCell         As Range
    Dim parentOptionCell   As Range
    Dim subOptionCell      As Range
    Dim updatedValue       As String
    Dim updatedOptionLabel As String
    Dim parentOptionLabel  As String
    Dim subOptionLabel     As String
    Dim defaultColorCode   As String
    
    For Each targetCell In updatedRange
        With targetCell
            Set parentOptionCell = .Offset(-1, 0)
            Set subOptionCell = .Offset(1, 0)
            
            updatedOptionLabel = .Offset(0, -1).Value
            parentOptionLabel = .Offset(-1, -1).Value
            subOptionLabel = .Offset(1, -1).Value

            updatedValue = .Value
            
            If updatedValue = vbNullString Then
                WriteNewRangeValue targetCell, GetDefaultCertificateOptions(updatedOptionLabel)
            End If

            Select Case parentOptionLabel
                Case "Type:", "Layout:", "Design:"
                    UpdateCertificateSubOptions subOptionCell, subOptionLabel, updatedValue
                Case "Border:"
                    If subOptionCell.Value = "Default" Then
                        
                        defaultColorCode = GetCertificateBorderColorCode(updatedValue, "Default")
                        
                        WriteNewRangeValue Options.Range("K16"), defaultColorCode
                        WriteNewRangeValue subOptionCell, GetCertificateBorderColorLabel(defaultColorCode)
                    End If

                    UpdateCertificateSubOptions subOptionCell, subOptionLabel, updatedValue
                Case "Border Color:"
                    WriteNewRangeValue subOptionCell, GetCertificateBorderColorCode(parentOptionCell.Value, updatedValue)

                    Select Case updatedValue
                        Case "Custom"
                            subOptionCell.Select
                        Case "Default"
                            WriteNewRangeValue targetCell, GetCertificateBorderColorLabel(subOptionCell.Value)
                    End Select
                Case "Color Code:"
                    updatedValue = NormalizeColorCode(updatedValue)
                    WriteNewRangeValue targetCell, updatedValue

                    If Not IsColorCodeValid(updatedValue) Then
                        WriteNewRangeValue targetCell, GetDefaultCertificateOptions(parentOptionLabel)
                    End If

                    CheckIfCustomBorderColorCode updatedValue, parentOptionCell
            End Select
        End With
    Next targetCell
End Sub

Private Sub UpdateCertificateSubOptions(ByVal subOptionCell As Range, ByRef subOptionLabel As String, ByVal parentValue As String)
    Dim currentFormula  As String
    Dim newOptionsList  As String
    Dim childCategory   As String
    Dim validationType  As Long
    
    If Trim$(parentValue) = vbNullString Then
        Exit Sub
    End If
    
    newOptionsList = GetCertificateValidationFormula(subOptionLabel, parentValue)
    
    With subOptionCell
        If g_UserOptions.EnableLogging Then
            DebugAndLogging GetMsg("Debug.Worksheet.UpdatingCertificateOptions", .Offset(-1, -1).Value, parentValue)
        End If

        With .Validation
            On Error Resume Next
            validationType = .Type
            On Error GoTo 0
            
            If validationType = xlValidateList Then
                currentFormula = .Formula1
            Else
                currentFormula = vbNullString
            End If
            
            If StrComp(Trim$(currentFormula), Trim$(newOptionsList), vbTextCompare) <> 0 Then
                On Error Resume Next
                .Delete
                On Error GoTo 0
                
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=newOptionsList
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End If
        End With
            
        If Not IsCurrentValueInNewList(.Value, newOptionsList) Then
            WriteNewRangeValue subOptionCell, GetDefaultCertificateOptions(subOptionLabel)
        End If
        
        childCategory = GetNextCertificateCategory(subOptionLabel)
        If childCategory <> vbNullString Then
            UpdateCertificateSubOptions .Offset(1, 0), childCategory, .Value
        End If
    End With
End Sub

Private Sub UpdateCertificatePreview()
    Dim certificateSettings As CertificateDesign
    Dim previewDesign As String
    Dim previewBorder As String
    Dim currentShapeName As String
    Dim currentShapeNameParts As Variant
    Dim shpVisible As Boolean
    Dim shp As Shape
    
    certificateSettings = LoadCertificateDesign()

    previewDesign = GenerateCertificatePreviewDesign(certificateSettings)
    previewBorder = GenerateCertificatePreviewBorder(certificateSettings)

    For Each shp In Options.Shapes
        With shp
            currentShapeName = .Name
            currentShapeNameParts = Split(currentShapeName, "_")
            
            If UBound(currentShapeNameParts) >= 1 Then
                Select Case currentShapeNameParts(1)
                    Case "Speech", "Winter", "Spring", "Summer", "Fall", "Autumn"
                        .Visible = (currentShapeName = previewDesign)
                    Case "Border"
                        If certificateSettings.borderType = "Disabled" Then
                            .Visible = msoFalse
                        Else
                            shpVisible = (currentShapeName = previewBorder)
                            .Visible = shpVisible
                            
                            If shpVisible Then
                                If ShapeHasFill(shp) Then
                                    shp.Fill.ForeColor.RGB = ConvertHexToRGB(certificateSettings.borderColorCode)
                                End If
                            End If
                        End If
                End Select
            End If
        End With
    Next shp
End Sub

Private Function ShapeHasFill(ByRef shp As Shape) As Boolean
    On Error Resume Next
    ShapeHasFill = shp.Fill.Visible
    On Error GoTo 0
End Function

Private Function GenerateCertificatePreviewDesign(ByRef certificateSettings As CertificateDesign) As String
    With certificateSettings
        GenerateCertificatePreviewDesign = "Layout_" & .Type & "_" & .Layout & "_" & .Design
    End With
End Function

Private Function GenerateCertificatePreviewBorder(ByRef certificateSettings As CertificateDesign) As String
    With certificateSettings
        GenerateCertificatePreviewBorder = "Embedded_Border_" & .Layout & "_" & .borderType
    End With
End Function

Private Function NormalizeColorCode(ByVal colorCode As String) As String
    If Left$(colorCode, 1) <> "#" Then colorCode = "#" & colorCode
    NormalizeColorCode = UCase$(colorCode)
End Function

Public Function ConvertHexToRGB(ByVal hexCode As String) As Long
    Dim redValue    As Long: redValue = CLng("&H" & Mid$(hexCode, 2, 2))
    Dim greenValue  As Long: greenValue = CLng("&H" & Mid$(hexCode, 4, 2))
    Dim blueValue   As Long: blueValue = CLng("&H" & Mid$(hexCode, 6, 2))
    
    ConvertHexToRGB = RGB(redValue, greenValue, blueValue)
End Function

Private Function GetNextCertificateCategory(ByVal currentCategory As String) As String
    Select Case currentCategory
        Case "Type:"
            GetNextCertificateCategory = "Layout:"
        Case "Layout:"
            GetNextCertificateCategory = "Design:"
        Case "Design:"
            GetNextCertificateCategory = "Border:"
        Case "Border:"
            GetNextCertificateCategory = "Border Color:"
        Case Else
            GetNextCertificateCategory = vbNullString
    End Select
End Function

Public Function GetDefaultCertificateOptions(ByVal optionLabel As String) As String
    Select Case optionLabel
        Case "Type:"
            GetDefaultCertificateOptions = "Speech Contest"
        Case "Layout:"
            GetDefaultCertificateOptions = "Landscape"
        Case "Design:"
            GetDefaultCertificateOptions = "Default"
        Case "Border:"
            GetDefaultCertificateOptions = "Disabled"
        Case "Border Color:"
            GetDefaultCertificateOptions = "Default"
        Case "Color Code:"
            GetDefaultCertificateOptions = GetCertificateBorderColorCode(Options.Range("K14").Value, Options.Range("K15").Value)
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

Private Sub CheckIfCustomBorderColorCode(ByVal newColorCode As String, ByVal borderColorSelectionValue As Range)
    Dim currentLabel As String
    Dim expectedLabel As String
    
    currentLabel = Trim$(borderColorSelectionValue.Value)
    expectedLabel = GetCertificateBorderColorLabel(UCase$(newColorCode))
    
    If currentLabel <> expectedLabel Then
        WriteNewRangeValue borderColorSelectionValue, expectedLabel
    End If
End Sub

Private Function IsColorCodeValid(ByVal colorCode As String) As Boolean
    Dim currentChar As String
    Dim i As Long
    
    If Left$(colorCode, 1) = "#" Then
        colorCode = Mid$(colorCode, 2)
    End If
    
    If Len(colorCode) <> 6 Then
        IsColorCodeValid = False
        Exit Function
    End If
    
    colorCode = UCase$(colorCode)
    
    For i = 1 To Len(colorCode)
        currentChar = Mid$(colorCode, i, 1)
        Select Case currentChar
            Case "0" To "9", "A" To "F"
                ' Valid character, so skip
            Case Else
                IsColorCodeValid = False
                Exit Function
        End Select
    Next i
    
    IsColorCodeValid = True
End Function

Private Function IsCurrentValueInNewList(ByVal currentValue As String, ByVal validationList As String) As Boolean
    Dim validationValues As Variant
    Dim matchFound As Boolean
    Dim i As Long
    
    matchFound = False
    validationValues = Split(validationList, ",")
    
    For i = LBound(validationValues) To UBound(validationValues)
        If currentValue = validationValues(i) Then
            matchFound = True
            Exit For
        End If
    Next i
    
    IsCurrentValueInNewList = matchFound
End Function

Public Function GetCertificateValidationFormula(ByVal targetCategory As String, ByVal parentCategory As String) As String
    Static layoutOptionsList      As New Dictionary
    Static designOptionsList      As New Dictionary
    Static borderOptionsList      As New Dictionary
    Static borderColorOptionsList As New Dictionary

    If Not layoutOptionsList.Exists("Test") Then
        layoutOptionsList.Add "Test", "Test"
        layoutOptionsList.Add "Speech Contest", "Landscape"
        layoutOptionsList.Add "Winter Speeches", "Landscape"
        layoutOptionsList.Add "Spring Speeches", "Landscape"
        layoutOptionsList.Add "Summer Speeches", "Landscape"
        layoutOptionsList.Add "Autumn Speeches", "Landscape"
        layoutOptionsList.Add "Fall Speeches", "Landscape"
    End If

    If Not designOptionsList.Exists("Test") Then
        designOptionsList.Add "Test", "Test"
        designOptionsList.Add "Landscape", "Default,Modern"
        designOptionsList.Add "Portrait", "Default,Modern"
    End If

    If Not borderOptionsList.Exists("Test") Then
        borderOptionsList.Add "Test", "Test"
        borderOptionsList.Add "Default", "Disabled,Style 1,Style 2"
        borderOptionsList.Add "Modern", "Disabled,Style 1,Style 2"
    End If
    
    If Not borderColorOptionsList.Exists("Test") Then
        borderColorOptionsList.Add "Test", "Test"
        borderColorOptionsList.Add "Style 1", "Default,Gold,Metalic Gold,Silver,Dark Teal,Custom"
        borderColorOptionsList.Add "Style 2", "Default,Gold,Metalic Gold,Silver,Dark Teal,Custom"
    End If

    On Error Resume Next
    Select Case targetCategory
        Case "Layout:"
            If layoutOptionsList.Exists(parentCategory) Then
                GetCertificateValidationFormula = layoutOptionsList(parentCategory)
            End If
        Case "Design:"
            If designOptionsList.Exists(parentCategory) Then
                GetCertificateValidationFormula = designOptionsList(parentCategory)
            End If
        Case "Border:"
            If borderOptionsList.Exists(parentCategory) Then
                GetCertificateValidationFormula = borderOptionsList(parentCategory)
            End If
        Case "Border Color:"
            If borderColorOptionsList.Exists(parentCategory) Then
                GetCertificateValidationFormula = borderColorOptionsList(parentCategory)
            End If
    End Select
    On Error GoTo 0
End Function