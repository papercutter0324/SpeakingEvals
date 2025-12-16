Option Explicit

Private Sub Worksheet_Change(ByVal modifiedCells As Range)
    Dim startTime As Date
    Dim currentCell As Range
    
    GetRunTime "Start", startTime

    ReadOptionsValues
    ToggleApplicationFeatures "Disable"
    ToggleSheetProtection Options, False

    If Not VerifyDictionariesAreLoaded Then
        InitializeDictionaries
    End If

    On Error GoTo ErrorHandler
    For Each currentCell In modifiedCells
        Select Case True
            Case Not Intersect(currentCell, Options.Range("K2")) Is Nothing
                g_UserOptions.DisplayEntryTips = IIf(Options.Range("K2").Value = "Yes", True, False)
                ToggleValidationTips
            Case Not Intersect(currentCell, Options.Range("K11:K16")) Is Nothing
                UpdateCertificateDesign Me, currentCell
        End Select
    Next currentCell
    On Error GoTo 0
    
Cleanup:
    ToggleApplicationFeatures "Enable"
    ToggleSheetProtection Options, True
    GetRunTime "End", startTime
    Exit Sub

ErrorHandler:
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.Worksheet.WorksheetError", "Worksheet_Change", Err.Number, Err.Description)
    End If
    Resume Cleanup
End Sub