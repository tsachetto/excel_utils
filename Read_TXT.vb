Sub readtxt()

Dim myFile As String, textline As String

myFile = "C:\Py\CNPJs_Table\Fornec.txt"

  Open myFile For Input As #1

    Do Until EOF(1)
    
      Line Input #1, textline
      Range("A200000").End(xlUp).Offset(1, 0) = textline

    Loop

  Close #1

End Sub
