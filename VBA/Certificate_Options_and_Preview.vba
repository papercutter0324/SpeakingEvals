Option Explicit

#Const PRINT_DEBUG_MESSAGES = True
#If Mac Then
    Const APPLE_SCRIPT_FILE As String = "SpeakingEvals.scpt"
    Const APPLE_SCRIPT_SPLIT_KEY = "-,-"
#End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Certificate Layout Updates
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateCertificateDesign(ByVal ws As Worksheet, ByVal updatedCellsRange As Range)
    Dim certificateSettings As Variant
    
    ToggleApplicationFeatures False
    ToggleSheetProtection ws, False
    On Error GoTo ErrorHandler
    
    ' Step 1: Update list options
    UpdateCertificateSettingsLists updatedCellsRange
    
    ' Step 2: Update the preview
    certificateSettings = ws.Range("J10:K14").Value
    UpdateCertificatePreview ws, certificateSettings
    
CleanUp:
    ToggleApplicationFeatures True
    ToggleSheetProtection ws, True
    Exit Sub
ErrorHandler:
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Error in Worksheet_Change: " & Err.Description & " (Error " & Err.Number & ")"
    #End If
    Resume CleanUp
End Sub

Private Sub UpdateCertificatePreview(ByVal ws As Worksheet, ByVal certificateSettings As Variant)
    Dim previewDesign As String
    Dim previewBorder As String
    Dim previewShapeName As String
    Dim shapeColor As Long
    Dim shp As Shape
    Dim shtShapes As Shapes
    
    Const PREVIEW_DESIGN_PREFIX As String = "Preview_Design_"
    Const PREVIEW_BORDER_PREFIX As String = "Preview_Border_"
    
    Set shtShapes = ws.Shapes
    
    previewDesign = "Preview_Design_" & certificateSettings(1, 2) & "_" & certificateSettings(2, 2)
    previewBorder = "Preview_Border_" & certificateSettings(1, 2) & "_" & certificateSettings(3, 2)

    ' Step 1: Toggle correct layout
    For Each shp In shtShapes
        With shp
            previewShapeName = .Name
            Select Case True
                Case (PREVIEW_DESIGN_PREFIX = Left$(previewShapeName, (Len(PREVIEW_DESIGN_PREFIX))))
                    .Visible = (.Name = previewDesign)
                Case (PREVIEW_BORDER_PREFIX = Left$(previewShapeName, (Len(PREVIEW_DESIGN_PREFIX))))
                    If (certificateSettings(3, 2) = "Disabled") Then
                        .Visible = msoFalse
                    Else
                        .Visible = (.Name = previewBorder)
                        
                        If (.Name = previewBorder) Then
                            shapeColor = ConvertHexToRGB(certificateSettings(5, 2))
                            If shp.Fill.ForeColor.RGB <> shapeColor Then
                                shp.Fill.ForeColor.RGB = shapeColor
                            End If
                        End If
                    End If
            End Select
        End With
    Next shp
End Sub

Public Function ConvertHexToRGB(ByVal hexCode As String) As Long
    Dim redValue As Long
    Dim greenValue As Long
    Dim blueValue As Long
    
    redValue = CLng("&H" & Mid$(hexCode, 2, 2))
    greenValue = CLng("&H" & Mid$(hexCode, 4, 2))
    blueValue = CLng("&H" & Mid$(hexCode, 6, 2))
    
    ConvertHexToRGB = RGB(redValue, greenValue, blueValue)
End Function

Private Sub UpdateCertificateSettingsLists(ByVal updatedCellsRange As Range)
    Dim updatedCell As Range
    Dim targetValidationRange As Range
    Dim updatedValue As String
    Dim updatedCategory As String
    Dim targetCategory As String
    Dim colorCode As String
    Dim colorLabel As String
    
    For Each updatedCell In updatedCellsRange
        With updatedCell
            updatedValue = .Value
            updatedCategory = .Offset(0, -1).Value
            targetCategory = .Offset(1, -1).Value
            Set targetValidationRange = .Offset(1, 0)
        
            If updatedValue = vbNullString Then
                updatedValue = GetDefaultCertificateOptions(.Offset(0, -1), updatedCategory)
                .Value = updatedValue
            End If

            Select Case updatedCategory
                Case "Layout:", "Design:"
                    UpdateCertificateSubOptions targetValidationRange, targetCategory, updatedValue
                Case "Border:"
                    If .Offset(1, 0).Value = "Default" Then
                        colorCode = GetCertificateBorderColorCode(updatedValue, "Default")
                        colorLabel = GetCertificateBorderColorLabel(colorCode)
                        
                        .Offset(2, 0).Value = colorCode
                        .Offset(1, 0).Value = colorLabel
                    End If
                Case "Border Color:"
                    colorCode = GetCertificateBorderColorCode(.Offset(-1, 0).Value, updatedValue)
                    .Offset(1, 0).Value = colorCode
                    
                    Select Case updatedValue
                        Case "Custom"
                            .Offset(1, 0).Select
                        Case "Default"
                            .Value = GetCertificateBorderColorLabel(colorCode)
                    End Select
                Case "Color Code:"
                    If updatedValue <> UCase$(updatedValue) Then
                        updatedValue = UCase$(updatedValue)
                        .Value = updatedValue
                    End If
                    
                    If Left$(updatedValue, 1) <> "#" Then
                        updatedValue = "#" & updatedValue
                        .Value = updatedValue
                    End If
                    
                    If Not IsColorCodeValid(updatedValue) Then
                        .Value = GetDefaultCertificateOptions(updatedCell, updatedCategory)
                    End If
                    
                    CheckIfCustomBorderColorCode updatedValue, .Offset(-1, 0)
            End Select
        End With
    Next updatedCell
End Sub

Private Sub CheckIfCustomBorderColorCode(ByVal newColorCode As String, ByVal borderColorSelectionValue As Range)
    Dim currentLabel As String
    Dim expectedLabel As String
    
    currentLabel = Trim$(borderColorSelectionValue.Value)
    expectedLabel = GetCertificateBorderColorLabel(UCase$(newColorCode))
    
    If currentLabel <> expectedLabel Then
        borderColorSelectionValue.Value = expectedLabel
    End If
End Sub

Private Sub UpdateCertificateSubOptions(ByVal targetValidationRange As Range, ByRef targetCategory As String, ByVal updatedValue As String)
    Dim currentValidationFormula As String
    Dim updateValidationFormula As String
    
    #If PRINT_DEBUG_MESSAGES Then
        Debug.Print "Updating Certicate Lists" & vbNewLine & _
                    INDENT_LEVEL_1 & targetCategory & vbNewLine & _
                    INDENT_LEVEL_1 & updatedValue
    #End If
    
    ' Step 1: Update validation formula
    ' Update this to point at the dependant category
    currentValidationFormula = targetValidationRange.Validation.Formula1
    updateValidationFormula = GetCertificateValidationFormula(targetCategory, updatedValue)
    
    If currentValidationFormula <> updateValidationFormula Then
        On Error Resume Next
        targetValidationRange.Validation.Delete
        On Error GoTo 0
        
        With targetValidationRange.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=updateValidationFormula
            .InputTitle = vbNullString
            .InputMessage = vbNullString
            .ShowInput = True
            .ShowError = True
        End With
    End If
    
    If Not IsCurrentValueInNewList(targetValidationRange.Value, updateValidationFormula) Then
        targetValidationRange.Value = GetDefaultCertificateOptions(targetValidationRange, targetCategory)
    End If
    
    If targetCategory = "Design:" Then
        targetCategory = "Border:"
        UpdateCertificateSubOptions targetValidationRange.Offset(1, 0), targetCategory, targetValidationRange.Value
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
                ' Valid character
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

Public Function GetCertificateValidationFormula(ByVal validationCategory As String, ByVal validationLabel As String) As String
    If validationCategory = "Design:" Then
        Select Case validationLabel
            Case "Landscape"
                GetCertificateValidationFormula = "Default"
            Case "Portrait"
                GetCertificateValidationFormula = "Default"
        End Select
    ElseIf validationCategory = "Border:" Then
        Select Case validationLabel
            Case "Default"
                GetCertificateValidationFormula = "Disabled,Style 1,Style 2"
            Case "Modern"
                GetCertificateValidationFormula = "Disabled,Style 1,Style 2"
        End Select
    End If
End Function
