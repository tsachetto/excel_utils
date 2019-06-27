Sub MergePDFs()

'Antes de tudo, na janela do VBA > Referências e habilitar o Acrobat.
'Pronto, vamos nessa!

  ' --> Configs
  Const MyPath = "C:\Users\tsach\Desktop\Renomeador PDF"            ' Path where PDF files are stored
   Const MyFiles = "00001 - 1 .pdf,00001 - 2 .pdf"  ' List of PDFs to ne merged
  Const DestFile = "MergedFile.pdf"   ' The name of the merged file
  ' <-- Fim Configs
 
  Dim a As Variant, i As Long, n As Long, ni As Long, p As String
  Dim AcroApp As New Acrobat.AcroApp, PartDocs() As Acrobat.CAcroPDDoc
 
  If Right(MyPath, 1) = "\" Then p = MyPath Else p = MyPath & "\"
  a = Split(MyFiles, ",")
  ReDim PartDocs(0 To UBound(a))
 
  On Error GoTo exit_
  If Len(Dir(p & DestFile)) Then Kill p & DestFile
  For i = 0 To UBound(a)
    ' Verificar arquivo PDF
    If Dir(p & Trim(a(i))) = "" Then
      MsgBox "File not found" & vbLf & p & a(i), vbExclamation, "Canceled"
      Exit For
    End If
    ' Abrir PDF
    Set PartDocs(i) = CreateObject("AcroExch.PDDoc")
    PartDocs(i).Open p & Trim(a(i))
    If i Then
      ' Juntar arquivos PDF
      ni = PartDocs(i).GetNumPages()
      If Not PartDocs(0).InsertPages(n - 1, PartDocs(i), 0, ni, True) Then
        MsgBox "Cannot insert pages of" & vbLf & p & a(i), vbExclamation, "Canceled"
      End If
      ' Calcular numero de Páginas
      n = n + ni
      ' Liberar memória
      PartDocs(i).Close
      Set PartDocs(i) = Nothing
    Else
      ' Calcular numero de Páginas
      n = PartDocs(0).GetNumPages()
    End If
  Next
 
  If i > UBound(a) Then
    ' Salvar arquivo unido no destno
    If Not PartDocs(0).Save(PDSaveFull, p & DestFile) Then
      MsgBox "Cannot save the resulting document" & vbLf & p & DestFile, vbExclamation, "Canceled"
    End If
  End If
 
exit_:
 
  ' Erros
  If Err Then
    MsgBox Err.Description, vbCritical, "Error #" & Err.Number
  ElseIf i > UBound(a) Then
    MsgBox "The resulting file is created:" & vbLf & p & DestFile, vbInformation, "Done"
  End If
 
  ' Liberar memória
  If Not PartDocs(0) Is Nothing Then PartDocs(0).Close
  Set PartDocs(0) = Nothing
 
  ' Liberar Acrobat
  AcroApp.Exit
  Set AcroApp = Nothing
 
End Sub
