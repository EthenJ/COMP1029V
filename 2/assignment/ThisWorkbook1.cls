Private Sub Worksheet_Change(ByVal Target As Range)
If Target.Address = Range("CoefficientA").Address Or _
   Target.Address = Range("CoefficientB").Address Or _
   Target.Address = Range("CoefficientC").Address Then
    ...VBA code executed when one of the cells is changed...
End If
End Sub
