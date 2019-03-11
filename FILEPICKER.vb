Dim objFileDialog As Office.FileDialog

    Set objFileDialog = Application.FileDialog(MsoFileDialogType.msoFileDialogFilePicker)
    
    With objFileDialog
        .AllowMultiSelect = False
        .ButtonName = "Escolher"
        .Title = "Escolher arquivo:"
        .Filters.Add "Arquivos Excel", "*.xlsx; *.xlsm; *.xls; *.xlsb", 1
    
    'teste  
         If (.Show > 0) Then
             End If
         If (.SelectedItems.Count > 0) Then
             Call MsgBox(.SelectedItems(1))
         End If
         
    End With
