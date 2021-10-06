VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
  Dim Age As Integer
  Age = InputBox("Please enter your age:")
  
  Dim FirstDigit, SecondDigit As Integer
  SecondDigit = Age Mod 10
  FirstDigit = Int(Age / 10)
  
  Dim AgeString As String
  AgeString = ""
  
  If FirstDigit = 1 Then
    If Age = 10 Then
      AgeString = "ten"
    ElseIf Age = 11 Then
      AgeString = "eleven"
    ElseIf Age = 12 Then
      AgeString = "twelve"
    ElseIf Age = 13 Then
      AgeString = "thirteen"
    ElseIf Age = 14 Then
      AgeString = "fourteen"
    ElseIf Age = 15 Then
      AgeString = "fifhteen"
    ElseIf Age = 16 Then
      AgeString = "sixteen"
    ElseIf Age = 17 Then
      AgeString = "seventeen"
    ElseIf Age = 18 Then
      AgeString = "eighteen"
    ElseIf Age = 19 Then
      AgeString = "nineteen"
    End If
  Else
    If FirstDigit = 2 Then
      AgeString = "twenty"
    ElseIf FirstDigit = 3 Then
      AgeString = "thirty"
    ElseIf FirstDigit = 4 Then
      AgeString = "fourty"
    ElseIf FirstDigit = 5 Then
      AgeString = "fifty"
    ElseIf FirstDigit = 6 Then
      AgeString = "sixty"
    ElseIf FirstDigit = 7 Then
      AgeString = "seventy"
    ElseIf FirstDigit = 8 Then
      AgeString = "eighty"
    ElseIf FirstDigit = 9 Then
      AgeString = "ninety"
    End If
  
    If Age > 10 Then
      AgeString = AgeString + " "
    End If
  
    If SecondDigit = 1 Then
      AgeString = AgeString + "one"
    ElseIf SecondDigit = 2 Then
      AgeString = AgeString + "two"
    ElseIf SecondDigit = 3 Then
      AgeString = AgeString + "three"
    ElseIf SecondDigit = 4 Then
      AgeString = AgeString + "four"
    ElseIf SecondDigit = 5 Then
      AgeString = AgeString + "five"
    ElseIf SecondDigit = 6 Then
      AgeString = AgeString + "six"
    ElseIf SecondDigit = 7 Then
      AgeString = AgeString + "seven"
    ElseIf SecondDigit = 8 Then
      AgeString = AgeString + "eight"
    ElseIf SecondDigit = 9 Then
      AgeString = AgeString + "nine"
    End If
  
  End If
  
  MsgBox AgeString
End Sub
