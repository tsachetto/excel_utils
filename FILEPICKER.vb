Dim myFile As String
Dim objFileDialog As Office.FileDialog

'GetFile
    Set objFileDialog = Application.FileDialog(MsoFileDialogType.msoFileDialogFilePicker)
    
    With objFileDialog
        .AllowMultiSelect = False
    .ButtonName = "Escolher Texto"
    .Title = "Escolher Arquivo Texto:"
    .Filters.Add "Importar Texto", "*.txt", 1
    
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
