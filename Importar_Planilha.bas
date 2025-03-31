Sub ajuste_Razão()

Dim myFile As String
Dim objFileDialog As Office.FileDialog

Dim wbThis As Workbook
Dim wbTarget As Workbook
Dim wsSource As Worksheet
Dim wsTarget As Worksheet

Dim ul As Long

If MsgBox("Confirma excluir os dados atuais da planilha para atualizar por novos?", vbYesNo + vbInformation, "Ajuste Razão!") = vbNo Then
Exit Sub
Else: End If

'GetFile
    Set objFileDialog = Application.FileDialog(MsoFileDialogType.msoFileDialogFilePicker)
    
    With objFileDialog
        .AllowMultiSelect = False
    .ButtonName = "Escolher Razão"
    .Title = "Escolher Razão:"
    .Filters.Add "Importar Razão", "*.xl*", 1
    
    'teste
         If (.Show > 0) Then
             Exit Sub
             End If
         If (.SelectedItems.Count > 0) Then
            myFile = (.SelectedItems(1))
         End If
            
         If myFile = "" Then
            Exit Sub
         Else: End If
         
    End With
    
    'Limpar
    Sheets("Razão").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Cells.Select
    Selection.Delete Shift:=xlUp
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select

    
    'Define a planilha atual como wbThis
    Set wbThis = ThisWorkbook
    'Define a aba Plano_de_Contas como destino
    Set wsTarget = wbThis.Sheets("Razão")
    
    'Abre o arquivo especificado em myFile
    Set wbTarget = Workbooks.Open(myFile)
    'Define a primeira aba como fonte
    Set wsSource = wbTarget.Sheets(1)
    
    'Verificar se é o relatório desejado...
    If Left(Range("A4"), 5) <> "Razão" Then
    MsgBox "Relatório infomado incorretamente, verifique!"
    wbTarget.Close SaveChanges:=False
    Exit Sub
    Else: End If
    
    'Copia todos os dados da primeira aba
    wsSource.Cells.Copy
    
    'Cola os dados na aba Plano_de_Contas
    wbThis.Activate
    wsTarget.Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'Limpa a área de transferência
    Application.CutCopyMode = False
    
    'Fecha o arquivo aberto sem salvar alterações
    wbTarget.Close SaveChanges:=False
    
    'Garante que a planilha original permaneça ativa
    wsTarget.Activate

    End Sub
