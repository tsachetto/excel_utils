Private Sub Workbook_AfterSave(ByVal Success As Boolean)

Dim tplan, tln, tcl As Long
Dim i, j, z As Long
Dim texto As String
Dim mypath As String

Application.DisplayAlerts = False

'Diretorio do DB
mypath = "DIRETORIO A SER SALVO O TXT\NOME.TXT"
'Fim diretorio DB

'-- Criar TXT / Verificar se existe

If Dir(mypath) = "" Then

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim Fileout As Object
    Set Fileout = fso.CreateTextFile(mypath, True, True)
    Fileout.Close
    
Else: End If

'-- Fim criar TXT

tplan = ThisWorkbook.Worksheets.Count

'--Escrita no TXT

Open mypath For Output As #1

For i = 1 To tplan

    tln = Sheets(i).Range("D50000").End(xlUp).Row + 1
    tcl = Sheets(i).Range("EA1").End(xlToLeft).Column + 1
    
  '  Print #1, "Competencia|" & Sheets(i).Name
                    
        For j = 3 To tln
            
            texto = ""
            
            For z = 4 To tcl
                
                texto = texto & Trim(Sheets(i).Cells(j, z).Text) & "|"
                                
            Next z
            
          Print #1, Sheets(i).Name & Sheets(i).Cells(j, 4) & "|" & texto
                            
        Next j
    
Next i

Close #1

'--Fim Escrita no TXT
Application.DisplayAlerts = True
End Sub
