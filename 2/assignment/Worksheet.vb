Private Sub Worksheet_Change(ByVal Target As Range)
'Done by JIANG, Yicheng 20760840

If Target.Address = Range("CoefficientA").Address Or _
   Target.Address = Range("CoefficientB").Address Or _
   Target.Address = Range("CoefficientC").Address Then
  Dim A,B,C,D As Double
  A = Range("CoefficientA").Value
  B = Range("CoefficientB").Value
  C = Range("CoefficientC").Value
  D = B * B - 4 * A * C
  If D > 0 Then
    Range("Solution1").Value = (-B + Math.Sqr(D)) / (2 * A)
    Range("Solution2").Value = (-B - Math.Sqr(D)) / (2 * A)
  ElseIf D = 0 Then
    Range("Solution1").Value = -B / (2 * A)
    Range("Solution2").Value = "Same as solution 1"
  Else
    Range("Solution1").Value = "No solution"
    Range("Solution1").Value = "No solution"
  End If
End If
End Sub