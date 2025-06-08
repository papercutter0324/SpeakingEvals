Option Explicit

Private Sub Worksheet_Change(ByVal targetCellsRange As Range)
    UpdateCertificateDesign Me, targetCellsRange
End Sub
