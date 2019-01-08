Public popChangedCells As Scripting.Dictionary
Public recordChangingCells As Boolean

Private Sub Worksheet_Change(ByVal Target As Range)
    If (recordChangingCells = True) Then
        If (popChangedCells.exists(Target.Address(0, 0)) = False) Then
            popChangedCells.Add Target.Address(0, 0), Target
            
        End If
    End If
End Sub