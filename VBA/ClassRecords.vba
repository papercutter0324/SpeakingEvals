Option Explicit

Private Sub Worksheet_Change(ByVal targetCellsRange As Range)
    UpdateClassRecords Me, targetCellsRange
End Sub
