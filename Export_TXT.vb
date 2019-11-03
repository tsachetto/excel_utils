Sub exportar()

Dim mypath As String
Dim ln As Long

ln = Sheets("????????").Range("A1000000").End(xlUp).Row

Application.DisplayAlerts = False

'Diretorio do DB
mypath = Sheets("???????").Range("???????")
'Fim diretorio DB

'Verificar se o arquivo existe

If Dir(mypath) = "" Then

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim Fileout As Object
    Set Fileout = fso.CreateTextFile(mypath, True, True)
    Fileout.Close
    
Else: End If

'Criar Texto
Open mypath For Output As #1

For i = 1 To ln
            
 Print #1, Sheets("?????????").Range("??????" & i)
                                
Next i

Close #1

'--Fim Escrita no TXT
Application.DisplayAlerts = True
End Sub
