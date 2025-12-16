Option Explicit

Private Sub Worksheet_Change(ByVal targetCellsRange As Range)
    Dim startTime As Date
    
    GetRunTime "Start", startTime
    
    ReadOptionsValues
    ToggleApplicationFeatures "Disable"
    ToggleSheetProtection Me, False
    
    If Not VerifyDictionariesAreLoaded Then
        InitializeDictionaries
    End If
    
    If g_UserOptions.EnableLogging Then
        DebugAndLogging GetMsg("Debug.CodeExecution.EntryPoint", Me.Name), True
    End If
    
    UpdateClassRecords Me, targetCellsRange
    
    ToggleSheetProtection Me, True
    ToggleApplicationFeatures "Enable"
    GetRunTime "End", startTime
End Sub
