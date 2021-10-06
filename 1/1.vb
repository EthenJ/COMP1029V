Sub Workbook_Open()
    Dim Input As Integer
    
    Name = InputBox("What is your name?", "Input your name", "David")
    
    MsgBox "Hello, " & Name & "!", , "Greeting"
End Sub