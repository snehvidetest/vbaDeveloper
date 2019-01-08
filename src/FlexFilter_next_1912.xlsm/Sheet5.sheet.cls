Public groChangedCells As Scripting.Dictionary
Private Sub Worksheet_Change(ByVal Target As Range)
    If (Sheet1.recordChangingCells = True) Then
        If (groChangedCells.exists(Target.Address(0, 0)) = False) Then
            groChangedCells.Add Target.Address(0, 0), Target
        End If
    End If
End Sub